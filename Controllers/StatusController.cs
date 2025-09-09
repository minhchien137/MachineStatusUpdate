using MachineStatusUpdate.Models;
using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using System.Drawing;
using ZXing;
using MachineStatusUpdate.Services;


namespace MachineStatusUpdate.Controllers
{
    public class StatusController : Controller
    {

        private readonly ApplicationDbContext _context;

        private readonly IWebHostEnvironment _webHostEnvironment;

        private readonly IStatusUpdateService _statusUpdateService;


        public StatusController(ApplicationDbContext context, IWebHostEnvironment webHostEnvironment, IStatusUpdateService statusUpdateService)
        {
            _context = context;
            _webHostEnvironment = webHostEnvironment;
            _statusUpdateService = statusUpdateService;

        }


        [HttpGet]
        public IActionResult Create()
        {
            return View();
        }


        // Hàm check mã máy có khớp không
        [HttpPost]
        public async Task<IActionResult> ValidateCode([FromBody] ValidateCodeRequest request)
        {
            try
            {
                if (string.IsNullOrEmpty(request.Code))
                {
                    return Json(new { exists = false });
                }

                var exists = await _context.sVN_Equipment_Machine_Info
                    .AnyAsync(x => x.SVNCode == request.Code);

                return Json(new { exists = exists });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error validating code: {ex.Message}");
                return Json(new { exists = false });
            }
        }



        // Hàm xác định Operation dựa trên Code
        private async Task<string> GetOperationFromCodeAsync(string code)
        {
            if (string.IsNullOrEmpty(code))
                return "";

            try
            {
                var machineInfo = await _context.sVN_Equipment_Machine_Info
                    .FirstOrDefaultAsync(x => x.SVNCode == code);

                return machineInfo?.Project ?? "";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting operation from code: {ex.Message}");
                return "";
            }
        }

        // Hàm decode mã QR từ upload
        [HttpPost]
        public async Task<IActionResult> DecodeQR(IFormFile qrImage)
        {
            if (qrImage == null || qrImage.Length == 0)
                return Json(new { success = false, message = "Chưa chọn ảnh!" });

            try
            {
                using var stream = qrImage.OpenReadStream();
                using var skBitmap = SkiaSharp.SKBitmap.Decode(stream);

                var reader = new ZXing.SkiaSharp.BarcodeReader();
                var result = reader.Decode(skBitmap);

                if (result != null)
                {
                    return Json(new { success = true, code = result.Text });
                }
                else
                {
                    return Json(new { success = false, message = "Không đọc được mã từ ảnh!" });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Lỗi xử lý: " + ex.Message });
            }
        }

        // Method để xử lý lại toàn bộ dữ liệu (nếu cần)
        public async Task<IActionResult> ProcessAllHistoryToDetail()
        {
            try
            {
                // Xóa toàn bộ dữ liệu cũ trong bảng Detail
                var existingDetails = _context.SVN_Equipment_Status_Update_Detail.ToList();
                _context.SVN_Equipment_Status_Update_Detail.RemoveRange(existingDetails);
                await _context.SaveChangesAsync();

                // Lấy tất cả records từ History, group by Code và xử lý
                var allHistoryRecords = await _context.SVN_Equipment_Info_History_Test
                    .OrderBy(x => x.Code)
                    .ThenBy(x => x.Datetime)
                    .ToListAsync();

                var groupedByCode = allHistoryRecords.GroupBy(x => x.Code).ToList();

                foreach (var codeGroup in groupedByCode)
                {
                    var records = codeGroup.OrderBy(x => x.Datetime).ToList();

                    for (int i = 0; i < records.Count; i++)
                    {
                        var currentRecord = records[i];
                        if (!currentRecord.Datetime.HasValue) continue;

                        // Tính EstimateTime (số phút từ EstimateTime - DateTime)
                        double? estimateTimeMinutes = null;
                        if (!string.IsNullOrEmpty(currentRecord.EstimateTime))
                        {
                            if (TimeSpan.TryParse(currentRecord.EstimateTime, out TimeSpan estimateTimeSpan))
                            {
                                var estimateDateTime = currentRecord.Datetime.Value.Date.Add(estimateTimeSpan);
                                var timeDifference = estimateDateTime - currentRecord.Datetime.Value;
                                estimateTimeMinutes = timeDifference.TotalMinutes;
                            }
                        }

                        // Xử lý ToTime và DurationMinutes
                        string toTime = "";
                        float durationMinutes = 0;

                        if (i < records.Count - 1)
                        {
                            var nextRecord = records[i + 1];
                            if (nextRecord.Datetime.HasValue)
                            {
                                durationMinutes = (float)(nextRecord.Datetime.Value - currentRecord.Datetime.Value).TotalMinutes;
                                toTime = Math.Round(durationMinutes, 2).ToString(); // ToTime là số phút
                            }
                        }

                        // Tạo record mới cho bảng Detail
                        var detailRecord = new SVN_Equipment_Status_Update_Detail
                        {
                            Name = currentRecord.Name ?? "",
                            Operation = currentRecord.Operation ?? "",
                            State = currentRecord.State ?? "",
                            EstimateTime = estimateTimeMinutes?.ToString("F2") ?? "",
                            FromTime = currentRecord.Datetime.Value.ToString("yyyy-MM-dd HH:mm:ss"),
                            ToTime = toTime, // Lưu số phút, để rỗng nếu chưa có
                            DurationMinutes = durationMinutes
                        };

                        _context.SVN_Equipment_Status_Update_Detail.Add(detailRecord);
                    }
                }

                await _context.SaveChangesAsync();

                return Json(new { success = true, message = "Đã xử lý thành công toàn bộ dữ liệu History vào Detail!" });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi trong ProcessAllHistoryToDetail: {ex.Message}");
                return Json(new { success = false, message = $"Lỗi: {ex.Message}" });
            }
        }



        [HttpPost]
        public async Task<IActionResult> Create(SVN_Equipment_Info_History_Test model, IFormFile imageFile)
        {
            try
            {
                if (string.IsNullOrEmpty(model.Code) || string.IsNullOrEmpty(model.State))
                {
                    return Json(new { success = false, message = "Vui lòng điền đầy đủ thông tin bắt buộc!" });
                }

                var machineExists = await _context.sVN_Equipment_Machine_Info
                    .AnyAsync(x => x.SVNCode == model.Code);

                if (!machineExists)
                {
                    return Json(new { success = false, message = "Không tồn tại mã máy này trong hệ thống!" });
                }

                string imagePath = null;
                if (imageFile != null && imageFile.Length > 0)
                {
                    var allowedExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp" };
                    var fileExtension = Path.GetExtension(imageFile.FileName).ToLower();

                    if (!allowedExtensions.Contains(fileExtension))
                    {
                        return Json(new { success = false, message = "Chỉ cho phép upload ảnh với định dạng: jpg, jpeg, png, gif, bmp" });
                    }
                    if (imageFile.Length > 5 * 1024 * 1024)
                    {
                        return Json(new { success = false, message = "Kích thước ảnh không được vượt quá 5MB" });
                    }

                    var uploadsFolder = Path.Combine(_webHostEnvironment.WebRootPath, "uploads", "status-images");
                    if (!Directory.Exists(uploadsFolder))
                    {
                        Directory.CreateDirectory(uploadsFolder);
                    }
                    var fileName = $"{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}{fileExtension}";
                    var filePath = Path.Combine(uploadsFolder, fileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await imageFile.CopyToAsync(stream);
                    }
                    imagePath = $"/uploads/status-images/{fileName}";
                }

                string generateName = model.Code;
                if (!string.IsNullOrEmpty(model.Code) && model.Code.Contains("-"))
                {
                    var parts = model.Code.Split('-');
                    if (parts.Length >= 2 && int.TryParse(parts[1], out int number))
                    {
                        generateName = $"#{number}";
                    }
                }

                model.Name = generateName;
                model.Operation = await GetOperationFromCodeAsync(model.Code);
                model.Datetime = DateTime.Now;

                int insertedId = 0;
                using (var command = _context.Database.GetDbConnection().CreateCommand())
                {
                    command.CommandText = "EXEC [dbo].[SVN_InsertMachineStatus_Test] @Code, @Name, @State, @Operation, @EstimateTime, @Description, @Image, @Datetime";
                    command.CommandType = System.Data.CommandType.Text;

                    command.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Code", model.Code ?? ""));
                    command.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Name", model.Name ?? ""));
                    command.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@State", model.State ?? ""));
                    command.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Operation", model.Operation ?? ""));
                    command.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@EstimateTime", model.EstimateTime ?? ""));
                    command.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Description", model.Description ?? ""));
                    command.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Image", imagePath ?? ""));
                    command.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Datetime", model.Datetime));

                    if (command.Connection.State != System.Data.ConnectionState.Open)
                        await command.Connection.OpenAsync();

                    var result = await command.ExecuteScalarAsync();
                    insertedId = Convert.ToInt32(result);
                }


                var insertedRecord = await _context.SVN_Equipment_Info_History_Test
                    .FirstOrDefaultAsync(x => x.Id == insertedId);

                await _statusUpdateService.ProcessSingleRecordToUpdateDetail(insertedRecord);

                return Json(new { success = true, message = "Lưu trạng thái thành công!", data = insertedRecord });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in Create: {ex.Message}");
                return Json(new { success = false, message = $"Có lỗi xảy ra: {ex.Message}" });
            }
        }

        // Method để xử lý dữ liệu từ Detail sang Status Update
        [HttpPost]
        public async Task<IActionResult> ProcessToStatusUpdate(DateTime? filterDate = null)
        {
            try
            {
                await _statusUpdateService.ProcessDataToStatusUpdate(filterDate);
                return Json(new { success = true, message = "Đã xử lý thành công dữ liệu vào bảng Status Update!" });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi trong ProcessToStatusUpdate: {ex.Message}");
                return Json(new { success = false, message = $"Lỗi: {ex.Message}" });
            }
        }

        // Method hiển thị Status Update Report
        public async Task<IActionResult> StatusUpdateReport(DateTime? filterDate = null, string operation = "", int page = 1, int pageSize = 25)
        {
            try
            {
                var query = _context.SVN_Equipment_Status_Update.AsQueryable();

                // Apply date filter
                if (filterDate.HasValue)
                {
                    query = query.Where(x => x.Datetime.Date == filterDate.Value.Date);
                }

                if (!string.IsNullOrEmpty(operation))
                {
                    query = query.Where(x => x.Operation.Contains(operation));
                }

                var totalRecords = await query.CountAsync();

                var results = await query
                    .OrderByDescending(x => x.Datetime)
                    .ThenBy(x => x.Name)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize)
                    .AsNoTracking()
                    .ToListAsync();

                var totalPages = (int)Math.Ceiling((double)totalRecords / pageSize);

                // Pagination ViewBag
                ViewBag.CurrentPage = page;
                ViewBag.TotalPages = totalPages;
                ViewBag.PageSize = pageSize;
                ViewBag.TotalRecords = totalRecords;
                ViewBag.HasPreviousPage = page > 1;
                ViewBag.HasNextPage = page < totalPages;

                // Filter ViewBag
                ViewBag.FilterDate = filterDate?.ToString("yyyy-MM-dd") ?? "";
                ViewBag.Operation = operation ?? "";

                return View(results);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"Lỗi: {ex.Message}";
                ViewBag.FilterDate = filterDate?.ToString("yyyy-MM-dd") ?? "";
                ViewBag.Operation = operation ?? "";

                // Set default pagination values for error case
                ViewBag.CurrentPage = 1;
                ViewBag.TotalPages = 0;
                ViewBag.PageSize = pageSize;
                ViewBag.TotalRecords = 0;
                ViewBag.HasPreviousPage = false;
                ViewBag.HasNextPage = false;

                return View(new List<SVN_Equipment_Status_Update>());
            }
        }


        // Method xuất Excel cho Status Update
        public async Task<IActionResult> ExportStatusUpdateToExcel(DateTime? filterDate = null, string operation = "")
        {
            try
            {
                var query = _context.SVN_Equipment_Status_Update.AsQueryable();

                if (filterDate.HasValue)
                {
                    query = query.Where(x => x.Datetime.Date == filterDate.Value.Date);
                }

                if (!string.IsNullOrEmpty(operation))
                {
                    query = query.Where(x => x.Operation.Contains(operation));
                }

                var data = await query
                    .OrderByDescending(x => x.Datetime)
                    .ThenBy(x => x.Name)
                    .ToListAsync();

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("StatusUpdateReport");
                    var currentRow = 1;

                    // Font mặc định
                    ws.Style.Font.FontName = "Times New Roman";
                    ws.Style.Font.FontSize = 11;

                    // Header
                    string[] headers = { "Id", "Name", "Operation", "Start Time", "Duration (min)", "Total Downtime (min)", "Date" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        var cell = ws.Cell(currentRow, i + 1);
                        cell.Value = headers[i];
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.5);
                        cell.Style.Font.FontColor = XLColor.White;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    }

                    // Data rows
                    foreach (var item in data)
                    {
                        currentRow++;
                        ws.Cell(currentRow, 1).Value = item.Id;
                        ws.Cell(currentRow, 2).Value = item.Name;
                        ws.Cell(currentRow, 3).Value = item.Operation;
                        ws.Cell(currentRow, 4).Value = item.StartTime;
                        ws.Cell(currentRow, 5).Value = Math.Round(item.Duration, 2);
                        ws.Cell(currentRow, 6).Value = Math.Round(item.TotalDuration, 2);
                        ws.Cell(currentRow, 7).Value = item.Datetime.ToString("yyyyMMdd");
                    }

                    // Styling
                    ws.Columns(1, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Columns(1, 7).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    // Column widths
                    ws.Column(1).Width = 8;   // Id
                    ws.Column(2).Width = 15;  // Name
                    ws.Column(3).Width = 20;  // Operation
                    ws.Column(4).Width = 20;  // Start Time
                    ws.Column(5).Width = 15;  // Duration
                    ws.Column(6).Width = 18;  // Total Downtime
                    ws.Column(7).Width = 12;  // Date

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return File(stream.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"StatusUpdateReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi ExportStatusUpdateToExcel: {ex.Message}");
                return Json(new { success = false, message = $"Lỗi xuất Excel: {ex.Message}" });
            }
        }

        // API endpoint để xử lý dữ liệu cho ngày cụ thể
        [HttpPost]
        public async Task<IActionResult> ProcessDataForDate([FromBody] ProcessDateRequest request)
        {
            try
            {
                DateTime? filterDate = null;
                if (!string.IsNullOrEmpty(request.Date))
                {
                    if (DateTime.TryParse(request.Date, out DateTime parsedDate))
                    {
                        filterDate = parsedDate;
                    }
                }

                await _statusUpdateService.ProcessDataToStatusUpdate(filterDate);
                return Json(new { success = true, message = "Đã xử lý thành công!" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi: {ex.Message}" });
            }
        }


        // Method hiển thị kết quả nhập trạng thái
        public async Task<IActionResult> Result(string code = "", string state = "", string operation = "",
            string fromInsDateTime = "", string toInsDateTime = "", int page = 1, int pageSize = 25)
        {
            try
            {
                var query = _context.SVN_Equipment_Info_History_Test.AsQueryable();

                // Apply filter

                if (!string.IsNullOrEmpty(code))
                    query = query.Where(x => x.Code.Contains(code));

                if (!string.IsNullOrEmpty(state))
                    query = query.Where(x => x.State.Contains(state));

                if (!string.IsNullOrEmpty(operation))
                    query = query.Where(x => x.Operation.Contains(operation));

                if (!string.IsNullOrEmpty(fromInsDateTime) && DateTime.TryParse(fromInsDateTime, out var fromDate))
                {
                    query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value.Date >= fromDate.Date);
                }

                if (!string.IsNullOrEmpty(toInsDateTime) && DateTime.TryParse(toInsDateTime, out var toDate))
                {
                    query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value.Date <= toDate.Date);
                }

                var totalRecords = await query.CountAsync();

                var results = await query
                    .OrderByDescending(x => x.Datetime)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize)
                    .AsNoTracking()
                    .ToListAsync();

                var totalPages = (int)Math.Ceiling((double)totalRecords / pageSize);

                ViewBag.CurrentPage = page;
                ViewBag.TotalPages = totalPages;
                ViewBag.PageSize = pageSize;
                ViewBag.TotalRecords = totalRecords;
                ViewBag.HasPreviousPage = page > 1;
                ViewBag.HasNextPage = page < totalPages;

                // Truyền giá trị filter ra View

                ViewBag.Code = code ?? "";
                ViewBag.State = state ?? "";
                ViewBag.Operation = operation ?? "";
                ViewBag.fromInsDateTime = fromInsDateTime ?? "";
                ViewBag.toInsDateTime = toInsDateTime ?? "";

                return View(results);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"Lỗi: {ex.Message}";
                ViewBag.Code = code ?? "";
                ViewBag.State = state ?? "";
                ViewBag.Operation = operation ?? "";
                ViewBag.fromInsDateTime = fromInsDateTime ?? "";
                ViewBag.toInsDateTime = toInsDateTime ?? "";

                // Set default pagination values for error case
                ViewBag.CurrentPage = 1;
                ViewBag.TotalPages = 0;
                ViewBag.PageSize = pageSize;
                ViewBag.TotalRecords = 0;
                ViewBag.HasPreviousPage = false;
                ViewBag.HasNextPage = false;

                return View(new List<SVN_Equipment_Info_History_Test>());
            }
        }



        // Xuất File Excel kết quả
        public async Task<IActionResult> ExportToExcel(string code = "", string state = "", string operation = "", string fromInsDateTime = "", string toInsDateTime = "")
        {
            var query = _context.SVN_Equipment_Info_History_Test.AsQueryable();

            if (!string.IsNullOrEmpty(code))
                query = query.Where(x => x.Code.Contains(code));

            if (!string.IsNullOrEmpty(state))
                query = query.Where(x => x.State.Contains(state));

            if (!string.IsNullOrEmpty(operation))
                query = query.Where(x => x.Operation.Contains(operation));

            if (!string.IsNullOrEmpty(fromInsDateTime) && DateTime.TryParse(fromInsDateTime, out var fromDate))
            {
                query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value.Date >= fromDate.Date);
            }

            if (!string.IsNullOrEmpty(toInsDateTime) && DateTime.TryParse(toInsDateTime, out var toDate))
            {
                query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value.Date <= toDate.Date);
            }

            // Sắp xếp bản ghi theo thời gian ASC
            var data = await query.OrderBy(x => x.Datetime).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("StatusHistory");
                var currentRow = 1;

                // Font mặc định
                ws.Style.Font.FontName = "Times New Roman";
                ws.Style.Font.FontSize = 11;

                // Header
                string[] headers = { "Id", "Code", "Name", "State", "Operation", "Description", "Image", "Datetime" };
                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = ws.Cell(currentRow, i + 1);
                    cell.Value = headers[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.5);
                    cell.Style.Font.FontColor = XLColor.White;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                }

                // Thiết lập chiều cao hàng cho data (để ảnh hiển thị đẹp)
                const double rowHeight = 70;

                foreach (var item in data)
                {
                    currentRow++;
                    ws.Row(currentRow).Height = rowHeight;
                    ws.Cell(currentRow, 1).Value = item.Id;
                    ws.Cell(currentRow, 2).Value = item.Code;
                    ws.Cell(currentRow, 3).Value = item.Name;
                    ws.Cell(currentRow, 4).Value = item.State;
                    ws.Cell(currentRow, 5).Value = item.Operation;
                    ws.Cell(currentRow, 6).Value = item.Description;

                    if (!string.IsNullOrEmpty(item.Image))
                    {
                        try
                        {
                            string imagePath = "";
                            if (item.Image.StartsWith("/uploads/"))
                            {
                                imagePath = Path.Combine(_webHostEnvironment.WebRootPath, item.Image.TrimStart('/').Replace('/', Path.DirectorySeparatorChar));
                            }
                            else
                            {
                                imagePath = item.Image;
                            }

                            if (System.IO.File.Exists(imagePath))
                            {

                                var picture = ws.AddPicture(imagePath);
                                picture.MoveTo(ws.Cell(currentRow, 7), 8, 5);
                                picture.WithSize(100, 70);


                                var imageCell = ws.Cell(currentRow, 7);
                                imageCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                imageCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            }
                            else
                            {

                                ws.Cell(currentRow, 7).Value = "No image";
                                ws.Cell(currentRow, 7).Style.Font.FontColor = XLColor.Gray;
                            }
                        }
                        catch (Exception ex)
                        {

                            ws.Cell(currentRow, 7).Value = $"Error: {ex.Message}";
                            ws.Cell(currentRow, 7).Style.Font.FontColor = XLColor.Red;
                        }
                    }
                    else
                    {
                        ws.Cell(currentRow, 7).Value = "No image";
                        ws.Cell(currentRow, 7).Style.Font.FontColor = XLColor.Gray;
                    }
                    ws.Cell(currentRow, 8).Value = item.Datetime?.ToString("yyyy-MM-dd HH:mm:ss");
                }

                // Canh giữa các cột số và ngày
                ws.Columns(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Columns(2, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Columns(3, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Columns(4, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Columns(5, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Columns(7, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Column(1).Width = 8;
                ws.Column(2).Width = 15;
                ws.Column(3).Width = 15;
                ws.Column(4).Width = 15;
                ws.Column(5).Width = 15;
                ws.Column(6).Width = 15;
                ws.Column(7).Width = 15;
                ws.Column(8).Width = 18;

                using (var stream = new MemoryStream())
                {

                    workbook.SaveAs(stream);
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "StatusHistory.xlsx");
                }

            }

        }


        public async Task<IActionResult> DowntimeDetailReport(
    string code = "",
     string state = "",
     string operation = "",
     string fromInsDateTime = "",
     string toInsDateTime = "",
     int page = 1,
     int pageSize = 25)
        {
            try
            {
                IQueryable<SVN_Equipment_Status_Update_Detail> query = _context.SVN_Equipment_Status_Update_Detail;

                // Lọc theo Code
                if (!string.IsNullOrEmpty(code))
                {
                    query = query.Where(x => x.Name.Contains(code));
                }

                // Lọc theo State
                if (!string.IsNullOrEmpty(state))
                {
                    query = query.Where(x => x.State.Contains(state));
                }

                // Lọc theo Operation
                if (!string.IsNullOrEmpty(operation))
                {
                    query = query.Where(x => x.Operation.Contains(operation));
                }

                // Lọc theo khoảng thời gian
                if (!string.IsNullOrEmpty(fromInsDateTime) && DateTime.TryParse(fromInsDateTime, out DateTime fromDate))
                {
                    query = query.Where(x => x.FromTime.CompareTo(fromDate.ToString("yyyy-MM-dd HH:mm:ss")) >= 0);
                }

                if (!string.IsNullOrEmpty(toInsDateTime) && DateTime.TryParse(toInsDateTime, out DateTime toDate))
                {
                    query = query.Where(x => x.FromTime.CompareTo(toDate.ToString("yyyy-MM-dd HH:mm:ss")) <= 0);
                }

                // Lấy tổng số bản ghi
                var totalRecords = await query.CountAsync();
                var totalPages = (int)Math.Ceiling(totalRecords / (double)pageSize);

                // Áp dụng phân trang
                var pagedResults = await query
                    .OrderByDescending(x => x.FromTime)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize)
                    .ToListAsync();

                // Gán dữ liệu vào ViewBag để truyền sang View
                ViewBag.Code = code;
                ViewBag.State = state;
                ViewBag.Operation = operation;
                ViewBag.fromInsDateTime = fromInsDateTime;
                ViewBag.toInsDateTime = toInsDateTime;
                ViewBag.CurrentPage = page;
                ViewBag.PageSize = pageSize;
                ViewBag.TotalPages = totalPages;
                ViewBag.TotalRecords = totalRecords;
                ViewBag.HasPreviousPage = page > 1;
                ViewBag.HasNextPage = page < totalPages;

                return View(pagedResults);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"Lỗi DowntimeDetailReport: {ex.Message}";
                return View(new List<SVN_Equipment_Status_Update_Detail>());
            }
        }

        // Xuất Excel cho Downtime Detail
        public async Task<IActionResult> ExportDowntimeDetailToExcel(
            string code = "",
            string state = "",
            string operation = "",
            string fromInsDateTime = "",
            string toInsDateTime = "")
        {
            try
            {
                IQueryable<SVN_Equipment_Status_Update_Detail> query = _context.SVN_Equipment_Status_Update_Detail;

                // Lọc theo Code
                if (!string.IsNullOrEmpty(code))
                    query = query.Where(x => x.Name.Contains(code));

                // Lọc theo State
                if (!string.IsNullOrEmpty(state))
                    query = query.Where(x => x.State.Contains(state));

                // Lọc theo Operation
                if (!string.IsNullOrEmpty(operation))
                    query = query.Where(x => x.Operation.Contains(operation));

                // Lọc theo khoảng thời gian
                if (!string.IsNullOrEmpty(fromInsDateTime) && DateTime.TryParse(fromInsDateTime, out DateTime fromDate))
                {
                    query = query.Where(x => x.FromTime.CompareTo(fromDate.ToString("yyyy-MM-dd HH:mm:ss")) >= 0);
                }

                if (!string.IsNullOrEmpty(toInsDateTime) && DateTime.TryParse(toInsDateTime, out DateTime toDate))
                {
                    query = query.Where(x => x.FromTime.CompareTo(toDate.ToString("yyyy-MM-dd HH:mm:ss")) <= 0);
                }

                var data = await query.OrderByDescending(x => x.FromTime).ToListAsync();

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("DowntimeDetail");
                    var currentRow = 1;

                    // Font mặc định
                    ws.Style.Font.FontName = "Times New Roman";
                    ws.Style.Font.FontSize = 11;

                    // Header
                    string[] headers = { "Id", "Name", "Operation", "State", "Estimate Time (min)", "From Time", "To Time (min)", "Duration (min)" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        var cell = ws.Cell(currentRow, i + 1);
                        cell.Value = headers[i];
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.5);
                        cell.Style.Font.FontColor = XLColor.White;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    }

                    // Data rows
                    // Data rows
                    foreach (var item in data)
                    {
                        currentRow++;
                        ws.Cell(currentRow, 1).Value = item.Id;
                        ws.Cell(currentRow, 2).Value = item.Name;
                        ws.Cell(currentRow, 3).Value = item.Operation;
                        ws.Cell(currentRow, 4).Value = item.State;

                        // Estimate Time
                        if (!string.IsNullOrEmpty(item.EstimateTime) && double.TryParse(item.EstimateTime, out double estimateMinutes))
                        {
                            ws.Cell(currentRow, 5).Value = Math.Round(estimateMinutes, 1);
                            ws.Cell(currentRow, 5).Style.NumberFormat.Format = "0.0";
                        }
                        else
                        {
                            ws.Cell(currentRow, 5).Value = ""; // để trống
                        }

                        // FromTime
                        if (!string.IsNullOrEmpty(item.FromTime))
                        {
                            ws.Cell(currentRow, 6).Value = item.FromTime;
                        }
                        else
                        {
                            ws.Cell(currentRow, 6).Value = "";
                        }

                        // ToTime
                        if (!string.IsNullOrEmpty(item.ToTime) && double.TryParse(item.ToTime, out double toTimeMinutes))
                        {
                            ws.Cell(currentRow, 7).Value = Math.Round(toTimeMinutes, 1);
                            ws.Cell(currentRow, 7).Style.NumberFormat.Format = "0.0";
                        }
                        else
                        {
                            ws.Cell(currentRow, 7).Value = "";
                        }

                        // Duration
                        if (item.DurationMinutes > 0)
                        {
                            ws.Cell(currentRow, 8).Value = Math.Round(item.DurationMinutes, 1);
                            ws.Cell(currentRow, 8).Style.NumberFormat.Format = "0.0";
                        }
                        else
                        {
                            ws.Cell(currentRow, 8).Value = "";
                        }
                    }
                    // Styling
                    ws.Columns().AdjustToContents();

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return File(stream.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"DowntimeDetail_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi ExportDowntimeDetailToExcel: {ex.Message}");
                return Json(new { success = false, message = $"Lỗi xuất Excel: {ex.Message}" });
            }
        }



        // Hàm test
        [HttpGet]
        public async Task<IActionResult> TestProcessAll()
        {
            try
            {
                return await ProcessAllHistoryToDetail();
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message, stackTrace = ex.StackTrace });
            }
        }


        // DTO class cho request
        public class ProcessDateRequest
        {
            public string Date { get; set; }
        }

        public class InsertedIdResult
        {
            public int InsertedId { get; set; }
        }


        public class ValidateCodeRequest
        {
            public string Code { get; set; }
        }
    }
}
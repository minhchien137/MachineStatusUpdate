using MachineStatusUpdate.Models;
using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;


namespace MachineStatusUpdate.Controllers
{
    public class StatusController : Controller
    {

        private readonly ApplicationDbContext _context;

        private readonly IWebHostEnvironment _webHostEnvironment;


        public StatusController(ApplicationDbContext context, IWebHostEnvironment webHostEnvironment)
        {
            _context = context;
            _webHostEnvironment = webHostEnvironment;
        }


        [HttpGet]
        public IActionResult Create()
        {
            return View();
        }

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



        // Method để xác định Operation dựa trên Code từ bảng SVN_Equipment_Machine_Info
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



        [HttpPost]
        public async Task<IActionResult> Create(SVN_Equipment_Info_History model, IFormFile imageFile)
        {
            try
            {
                if (string.IsNullOrEmpty(model.Code) ||
                    string.IsNullOrEmpty(model.State))
                {
                    return Json(new { success = false, message = "Vui lòng điền đầy đủ thông tin bắt buộc!" });
                }

                // Kiểm tra xem Code có tồn tại trong bảng SVN_Equipment_Machine_Info không
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
                    if (parts.Length >= 2)
                    {
                        if (int.TryParse(parts[1], out int number))
                        {
                            generateName = $"#{number}";
                        }
                    }
                }

                model.Name = generateName;
                model.Operation = await GetOperationFromCodeAsync(model.Code);
                model.Datetime = DateTime.Now;

                // Gọi proc 
                await _context.Database.ExecuteSqlRawAsync(
                "EXEC [dbo].[SVN_InsertMachineStatus] {0}, {1}, {2}, {3}, {4}, {5}, {6}",
                    model.Code ?? "",
                    model.Name ?? "",
                    model.State ?? "",
                    model.Operation ?? "",
                    model.Description ?? "",
                    imagePath ?? "",
                    model.Datetime);

                Console.WriteLine("Stored procedure executed successfully - Operation: {model.Operation}");
                return Json(new { success = true, message = "Lưu trạng thái thành công!", data = model });

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in Create: {ex.Message}");
                return Json(new { success = false, message = $"Có lỗi xảy ra: {ex.Message}" });
            }
        }

        public async Task<IActionResult> Result(string code = "", string state = "", string operation = "",
            string fromInsDateTime = "", string toInsDateTime = "", int page = 1, int pageSize = 25)
        {
            try
            {
                var query = _context.SVN_Equipment_Info_History.AsQueryable();

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

                return View(new List<SVN_Equipment_Info_History>());
            }
        }



        // Xuất File Excel kết quả
        public async Task<IActionResult> ExportToExcel(string code = "", string state = "", string operation = "", string fromInsDateTime = "", string toInsDateTime = "")
        {
            var query = _context.SVN_Equipment_Info_History.AsQueryable();

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

        // ============= DOWNTIME DETAIL REPORT =============
        public async Task<IActionResult> DowntimeDetailReport(string code = "", string operation = "", string state = "",
            string fromInsDateTime = "", string toInsDateTime = "")
        {
            try
            {
                var query = _context.SVN_Equipment_Info_History.AsQueryable();

                if (!string.IsNullOrEmpty(code))
                    query = query.Where(x => x.Code.Contains(code));

                if (!string.IsNullOrEmpty(operation))
                    query = query.Where(x => x.Operation.Contains(operation));

                if (!string.IsNullOrEmpty(state))
                    query = query.Where(x => x.State.Contains(state));

                if (!string.IsNullOrEmpty(fromInsDateTime) && DateTime.TryParse(fromInsDateTime, out var fromDate))
                    query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value >= fromDate);

                if (!string.IsNullOrEmpty(toInsDateTime) && DateTime.TryParse(toInsDateTime, out var toDate))
                    query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value <= toDate);

                var history = await query.OrderBy(x => x.Datetime).ToListAsync();
                var report = new List<DowntimeReportDto>();

                for (int i = 0; i < history.Count - 1; i++)
                {
                    var current = history[i];
                    var next = history[i + 1];
                    if (!current.Datetime.HasValue || !next.Datetime.HasValue) continue;

                    var duration = next.Datetime.Value - current.Datetime.Value;

                    report.Add(new DowntimeReportDto
                    {
                        Code = current.Code,
                        Operation = current.Operation,
                        State = current.State,
                        FromTime = current.Datetime.Value,
                        ToTime = next.Datetime.Value,
                        DurationMinutes = duration.TotalMinutes
                    });
                }

                // Truyền giá trị filter ra View

                ViewBag.Code = code ?? "";
                ViewBag.State = state ?? "";
                ViewBag.Operation = operation ?? "";
                ViewBag.fromInsDateTime = fromInsDateTime ?? "";
                ViewBag.toInsDateTime = toInsDateTime ?? "";


                return View(report);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"Lỗi DowntimeDetailReport: {ex.Message}";
                return View(new List<DowntimeReportDto>());
            }
        }


        // ============= DOWNTIME SUMMARY REPORT =============
        public async Task<IActionResult> DowntimeSummaryReport(string code = "", string operation = "", string state = "",
            string fromInsDateTime = "", string toInsDateTime = "")
        {
            try
            {
                var query = _context.SVN_Equipment_Info_History.AsQueryable();

                if (!string.IsNullOrEmpty(code))
                    query = query.Where(x => x.Code.Contains(code));

                if (!string.IsNullOrEmpty(operation))
                    query = query.Where(x => x.Operation.Contains(operation));

                if (!string.IsNullOrEmpty(state))
                    query = query.Where(x => x.State.Contains(state));

                if (!string.IsNullOrEmpty(fromInsDateTime) && DateTime.TryParse(fromInsDateTime, out var fromDate))
                    query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value >= fromDate);

                if (!string.IsNullOrEmpty(toInsDateTime) && DateTime.TryParse(toInsDateTime, out var toDate))
                    query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value <= toDate);

                var history = await query.OrderBy(x => x.Datetime).ToListAsync();
                var detail = new List<DowntimeReportDto>();

                for (int i = 0; i < history.Count - 1; i++)
                {
                    var current = history[i];
                    var next = history[i + 1];
                    if (!current.Datetime.HasValue || !next.Datetime.HasValue) continue;

                    var duration = next.Datetime.Value - current.Datetime.Value;

                    detail.Add(new DowntimeReportDto
                    {
                        Code = current.Code,
                        Operation = current.Operation,
                        State = current.State,
                        FromTime = current.Datetime.Value,
                        ToTime = next.Datetime.Value,
                        DurationMinutes = duration.TotalMinutes
                    });
                }

                var summary = detail
                    .GroupBy(r => new { r.Code, r.Operation, r.State })
                    .Select(g => new DowntimeSummaryDto
                    {
                        Code = g.Key.Code,
                        Operation = g.Key.Operation,
                        State = g.Key.State,
                        TotalMinutes = g.Sum(x => x.DurationMinutes)
                    })
                    .ToList();


                ViewBag.Code = code ?? "";
                ViewBag.State = state ?? "";
                ViewBag.Operation = operation ?? "";
                ViewBag.fromInsDateTime = fromInsDateTime ?? "";
                ViewBag.toInsDateTime = toInsDateTime ?? "";

                return View(summary);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"Lỗi DowntimeSummaryReport: {ex.Message}";
                return View(new List<DowntimeSummaryDto>());
            }
        }

        // Xuất Excel Báo cáo chi tiết thời gian DownTime
        public async Task<IActionResult> ExportDowntimeDetailToExcel(string code = "", string operation = "", string state = "",
            string fromInsDateTime = "", string toInsDateTime = "")
        {
            var query = _context.SVN_Equipment_Info_History.AsQueryable();

            if (!string.IsNullOrEmpty(code))
                query = query.Where(x => x.Code.Contains(code));

            if (!string.IsNullOrEmpty(operation))
                query = query.Where(x => x.Operation.Contains(operation));

            if (!string.IsNullOrEmpty(state))
                query = query.Where(x => x.State.Contains(state));

            if (!string.IsNullOrEmpty(fromInsDateTime) && DateTime.TryParse(fromInsDateTime, out var fromDate))
                query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value >= fromDate);

            if (!string.IsNullOrEmpty(toInsDateTime) && DateTime.TryParse(toInsDateTime, out var toDate))
                query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value <= toDate);

            var history = await query.OrderBy(x => x.Datetime).ToListAsync();
            var report = new List<DowntimeReportDto>();

            for (int i = 0; i < history.Count - 1; i++)
            {
                var current = history[i];
                var next = history[i + 1];
                if (!current.Datetime.HasValue || !next.Datetime.HasValue) continue;

                var duration = next.Datetime.Value - current.Datetime.Value;

                report.Add(new DowntimeReportDto
                {
                    Code = current.Code,
                    Operation = current.Operation,
                    State = current.State,
                    FromTime = current.Datetime.Value,
                    ToTime = next.Datetime.Value,
                    DurationMinutes = Math.Round(duration.TotalMinutes, 2)
                });
            }

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("DowntimeDetail");
                var currentRow = 1;

                // Font mặc định
                ws.Style.Font.FontName = "Times New Roman";
                ws.Style.Font.FontSize = 11;

                // Header
                string[] headers = { "Code", "Operation", "State", "From", "To", "Duration (minutes)" };
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

                // Thiết lập chiều cao hàng cho data
                const double rowHeight = 25;

                foreach (var item in report)
                {
                    currentRow++;
                    ws.Row(currentRow).Height = rowHeight;
                    ws.Cell(currentRow, 1).Value = item.Code;
                    ws.Cell(currentRow, 2).Value = item.Operation;
                    ws.Cell(currentRow, 3).Value = item.State;
                    ws.Cell(currentRow, 4).Value = item.FromTime.ToString("yyyy-MM-dd HH:mm:ss");
                    ws.Cell(currentRow, 5).Value = item.ToTime.ToString("yyyy-MM-dd HH:mm:ss");
                    ws.Cell(currentRow, 6).Value = item.DurationMinutes;
                }

                // Canh giữa các cột
                ws.Columns(1, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Columns(1, 6).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                // Thiết lập chiều rộng cột
                ws.Column(1).Width = 15; // Code
                ws.Column(2).Width = 15; // Operation
                ws.Column(3).Width = 15; // State
                ws.Column(4).Width = 20; // From
                ws.Column(5).Width = 20; // To
                ws.Column(6).Width = 18; // Duration

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "DowntimeDetailReport.xlsx");
                }
            }
        }


        // Xuất Excel Báo cáo tổng hợp thời gian DownTime
        public async Task<IActionResult> ExportDowntimeSummaryToExcel(string code = "", string operation = "", string state = "",
            string fromInsDateTime = "", string toInsDateTime = "")
        {
            var query = _context.SVN_Equipment_Info_History.AsQueryable();

            if (!string.IsNullOrEmpty(code))
                query = query.Where(x => x.Code.Contains(code));

            if (!string.IsNullOrEmpty(state))
                query = query.Where(x => x.State.Contains(state));

            if (!string.IsNullOrEmpty(operation))
                query = query.Where(x => x.Operation.Contains(operation));

            if (!string.IsNullOrEmpty(fromInsDateTime) && DateTime.TryParse(fromInsDateTime, out var fromDate))
                query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value >= fromDate);

            if (!string.IsNullOrEmpty(toInsDateTime) && DateTime.TryParse(toInsDateTime, out var toDate))
                query = query.Where(x => x.Datetime.HasValue && x.Datetime.Value <= toDate);

            var history = await query.OrderBy(x => x.Datetime).ToListAsync();
            var detail = new List<DowntimeReportDto>();

            for (int i = 0; i < history.Count - 1; i++)
            {
                var current = history[i];
                var next = history[i + 1];
                if (!current.Datetime.HasValue || !next.Datetime.HasValue) continue;

                var duration = next.Datetime.Value - current.Datetime.Value;

                detail.Add(new DowntimeReportDto
                {
                    Code = current.Code,
                    Operation = current.Operation,
                    State = current.State,
                    FromTime = current.Datetime.Value,
                    ToTime = next.Datetime.Value,
                    DurationMinutes = duration.TotalMinutes
                });
            }

            var summary = detail
                .GroupBy(r => new { r.Code, r.Operation, r.State })
                .Select(g => new DowntimeSummaryDto
                {
                    Code = g.Key.Code,
                    Operation = g.Key.Operation,
                    State = g.Key.State,
                    TotalMinutes = Math.Round(g.Sum(x => x.DurationMinutes), 2)
                })
                .ToList();

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("DowntimeSummary");
                var currentRow = 1;

                // Font mặc định
                ws.Style.Font.FontName = "Times New Roman";
                ws.Style.Font.FontSize = 11;

                // Header
                string[] headers = { "Code", "Operation", "State", "Total Minutes" };
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

                // Thiết lập chiều cao hàng cho data
                const double rowHeight = 25;

                foreach (var item in summary)
                {
                    currentRow++;
                    ws.Row(currentRow).Height = rowHeight;
                    ws.Cell(currentRow, 1).Value = item.Code;
                    ws.Cell(currentRow, 2).Value = item.Operation;
                    ws.Cell(currentRow, 3).Value = item.State;
                    ws.Cell(currentRow, 4).Value = item.TotalMinutes;
                }

                // Canh giữa các cột
                ws.Columns(1, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Columns(1, 4).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                // Thiết lập chiều rộng cột
                ws.Column(1).Width = 15; // Code
                ws.Column(2).Width = 15; // Operation
                ws.Column(3).Width = 15; // State
                ws.Column(4).Width = 18; // Total Minutes

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "DowntimeSummaryReport.xlsx");
                }
            }
        }

        // DTO class
        public class DowntimeReportDto
        {
            public string Code { get; set; }
            public string Operation { get; set; }
            public string State { get; set; }
            public DateTime FromTime { get; set; }
            public DateTime ToTime { get; set; }
            public double DurationMinutes { get; set; }
        }

        // DTO cho tổng hợp
        public class DowntimeSummaryDto
        {
            public string Code { get; set; }
            public string Operation { get; set; }
            public string State { get; set; }
            public double TotalMinutes { get; set; }
        }


        public class ValidateCodeRequest
        {
            public string Code { get; set; }
        }
    }
}

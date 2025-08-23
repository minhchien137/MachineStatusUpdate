using MachineStatusUpdate.Models;
using Microsoft.AspNetCore.Mvc;
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

                model.Name = model.Code;
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
    }
}

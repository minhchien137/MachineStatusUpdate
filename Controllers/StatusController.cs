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


        [HttpPost]
        public async Task<IActionResult> Create(SVN_Equipment_Info_History model, IFormFile imageFile)
        {
            try
            {
                if (string.IsNullOrEmpty(model.Name) ||
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

                    var uploadsFolder = Path.Combine(_webHostEnvironment.WebRootPath, "uploads", "defect-images");
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

                    imagePath = $"/uploads/defect-images/{fileName}";
                }

                // Gọi proc 
                await _context.Database.ExecuteSqlRawAsync(

                "EXEC [dbo].[SVN_InsertMachineStatus] {0}, {1}, {2}, {3}, {4}, {5}, {6}",
                    model.Code ?? "",
                    model.Name ?? "",
                    model.State ?? "",
                    model.Operation ?? "",
                    model.Description ?? "",
                    imagePath ?? "",
                    DateTime.Now);

                Console.WriteLine("Stored procedure executed successfully");
                return Json(new { success = true, message = "Lưu thông tin thành công!", data = model });

            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Có lỗi xảy ra: {ex.Message}" });
            }
        }
    }
}

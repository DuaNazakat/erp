using Microsoft.AspNetCore.Mvc;

namespace erp.Controllers
{

        [Route("api/[controller]")]
        [ApiController]
        public class UploadController : ControllerBase
        {
            private readonly string[] allowedExtensions = { ".xlsx", ".pdf", ".ppt", ".pptx" };
            private readonly string uploadFolder = Path.Combine(Directory.GetCurrentDirectory(), "Uploads");

            [HttpPost("file")]
            public async Task<IActionResult> UploadFile(IFormFile file)
            {
                if (file == null || file.Length == 0)
                    return BadRequest("No file uploaded.");

                var extension = Path.GetExtension(file.FileName).ToLowerInvariant();

                if (!allowedExtensions.Contains(extension))
                    return BadRequest("Invalid file type. Only Excel, PDF, and PPT files are allowed.");

                if (!Directory.Exists(uploadFolder))
                    Directory.CreateDirectory(uploadFolder);

                var uniqueFileName = $"{Guid.NewGuid()}{extension}";
                var filePath = Path.Combine(uploadFolder, uniqueFileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                return Ok(new { message = "File uploaded successfully", fileName = uniqueFileName });
            }
        }
    
}

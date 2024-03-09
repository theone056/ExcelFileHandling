using ExcelFileHandling.Core.Models;
using ExcelFileHandlingAPI.Helpers;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.XSSF.UserModel;

namespace ExcelFileHandlingAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelFileController : ControllerBase
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public ExcelFileController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        [HttpPost("Upload")]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> Upload(CancellationToken cancellationToken)
        {
            string[] permittedExtension = { ".xls", ".xlsx"  };
            if (Request.Form.Files.Count == 0) return NoContent();

            var file = Request.Form.Files[0];

            if(!permittedExtension.Contains(Path.GetExtension(file.FileName).ToLowerInvariant()))
            {
                throw new BadHttpRequestException("Invalid File");
            }

            var stringPath = filePath(file);
            var test = ExcelHelper.Extract<IMSModel>(stringPath);
            return Ok();
        }

        private string filePath(IFormFile formFile)
        {
            if (formFile.Length == 0) throw new BadHttpRequestException("File is Empty");

            var fileExtension = Path.GetExtension(formFile.FileName);

            var root = _webHostEnvironment.WebRootPath;
            if(string.IsNullOrWhiteSpace(root))
            {
                root = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            }

            var folderPath = Path.Combine(root, "uploads");

            if(!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            var fileName = $"{Guid.NewGuid()}.{fileExtension}";
            var filePath = Path.Combine(folderPath, fileName);
            using(var stream = new FileStream(filePath,FileMode.Create))
            {
               formFile.CopyTo(stream);
            }

            return filePath;
        }
    }
}

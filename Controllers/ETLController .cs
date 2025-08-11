

using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using WebApplication1.Models;
using WebApplication1.Utils;

public class ETLController : Controller
{
    private readonly ETLService _etlService;

    public ETLController()
    {
        // Poné tu cadena de conexión acá
        string connectionString = "Server=ARGOM\\SQLEXPRESS;Database=ETL;Trusted_Connection=True;";
        _etlService = new ETLService(connectionString);
    }

    [HttpGet]
    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public async Task<IActionResult> UploadExcel()
    {
        var file = Request.Form.Files[0];
        if (file == null || file.Length == 0)
        {
            ViewBag.Message = "No se seleccionó archivo.";
            return View("Index");
        }

        var filePath = Path.GetTempFileName();

        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }

        List<Persona> personas = _etlService.LeerExcel(filePath);
        _etlService.GuardarEnBase(personas);

        ViewBag.Message = $"{personas.Count} registros cargados correctamente.";
       // return View("Index");
        return View("~/Views/Home/Index.cshtml");

    }
}

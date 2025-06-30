using Lab13_FrankBernaola.Application.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace Lab13_FrankBernaola.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ExcelController : ControllerBase
{
    private readonly IExcelService _excelService;

    public ExcelController(IExcelService excelService)
    {
        _excelService = excelService;
    }

    [HttpGet("crear-ejemplo")]
    public IActionResult Crear()
    {
        _excelService.CrearExcelConDatos();
        return Ok(new { mensaje = "✅ Excel creado en wwwroot/archivo.xlsx" });
    }

    [HttpGet("modificar-ejemplo")]
    public IActionResult Modificar()
    {
        _excelService.SecondExample();
        return Ok("Archivo modificado.");
    }
    
    [HttpGet("crear-tabla")]
    public IActionResult CrearTabla()
    {
        _excelService.ThirdExample();
        return Ok("✅ Archivo con tabla creado en wwwroot/reportes/tabla.xlsx");
    }
    
    [HttpGet("crear-estilos")]
    public IActionResult CrearConEstilos()
    {
        _excelService.FourthExample();
        return Ok("✅ Archivo con estilos creado ");
    }
    
}
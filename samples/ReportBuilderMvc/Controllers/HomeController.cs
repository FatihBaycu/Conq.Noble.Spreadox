using System.Diagnostics;
using Conq.Noble.Spreadox;
using Microsoft.AspNetCore.Localization;
using Microsoft.AspNetCore.Mvc;
using ReportBuilderMvc.Models;
using ReportBuilderMvc.Services;

namespace ReportBuilderMvc.Controllers;

public class HomeController : Controller
{
    private readonly ReportGeneratorService _reportService;
    private readonly ILocalizationService _loc;
    private readonly IWebHostEnvironment _env;

    // Track loaded fonts for UI selection
    private static readonly List<string> AvailableFonts = new() { "Liberation Sans" };

    public HomeController(ReportGeneratorService reportService, ILocalizationService loc, IWebHostEnvironment env)
    {
        _reportService = reportService;
        _loc = loc;
        _env = env;
    }

    [HttpGet]
    public IActionResult Index()
    {
        var model = new ReportConfigViewModel
        {
            ReportTitle = _loc["DefaultReportTitle"],
            DataSeriesCount = 10,
            Layout = ChartLayout.SeparatePages,
            Format = OutputFormat.Pdf,
            Header = new HeaderConfig
            {
                Title = _loc["DefaultHeaderTitle"],
                Text = _loc["DefaultHeaderText"],
                FontSize = 12,
                Bold = true,
                TitleColor = "#005293"
            },
            Footer = new FooterConfig
            {
                Text = _loc["DefaultFooterText"],
                PageNumberFormat = _loc["DefaultPageNumberFormat"]
            },
            Cover = new CoverConfig
            {
                Company = _loc["DefaultCoverCompany"],
                Title = _loc["DefaultCoverTitle"],
                Provider = _loc["DefaultCoverProvider"],
                IncludeDate = true
            }
        };

        model.Items = _reportService.GenerateRandomData(model.DataSeriesCount);
        ViewBag.AvailableFonts = AvailableFonts;
        return View(model);
    }

    [HttpPost]
    public IActionResult SetLanguage(string culture, string returnUrl = "/")
    {
        Response.Cookies.Append(
            CookieRequestCultureProvider.DefaultCookieName,
            CookieRequestCultureProvider.MakeCookieValue(new RequestCulture(culture)),
            new CookieOptions
            {
                Expires = DateTimeOffset.UtcNow.AddYears(1),
                IsEssential = true,
                SameSite = SameSiteMode.Lax
            }
        );

        return LocalRedirect(string.IsNullOrEmpty(returnUrl) ? "/" : returnUrl);
    }

    [HttpPost]
    public IActionResult GenerateRandomData([FromBody] RandomDataRequest request)
    {
        int count = Math.Clamp(request.Count, 1, 30);
        var items = _reportService.GenerateRandomData(count);
        return Json(new { success = true, items });
    }

    [HttpPost]
    public async Task<IActionResult> UploadFont(IFormFile? fontFile, string? fontName)
    {
        if (fontFile == null || fontFile.Length == 0)
        {
            return Json(new { success = false, message = _loc["SelectFontAlert"] });
        }

        if (!fontFile.FileName.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase))
        {
            return Json(new { success = false, message = _loc["OnlyTtfSupported"] });
        }

        string familyName = string.IsNullOrWhiteSpace(fontName)
            ? Path.GetFileNameWithoutExtension(fontFile.FileName)
            : fontName.Trim();

        using var memoryStream = new MemoryStream();
        await fontFile.CopyToAsync(memoryStream);
        byte[] fontBytes = memoryStream.ToArray();

        try
        {
            PdfFontRegistry.Register(familyName, fontBytes);

            lock (AvailableFonts)
            {
                if (!AvailableFonts.Contains(familyName, StringComparer.OrdinalIgnoreCase))
                {
                    AvailableFonts.Add(familyName);
                }
            }

            return Json(new { success = true, familyName, message = _loc["FontRegisteredSuccess", familyName] });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = _loc["FontRegistrationFailed", ex.Message] });
        }
    }

    [HttpPost]
    public async Task<IActionResult> Export(
        [FromForm] ReportConfigViewModel model,
        IFormFile? headerLogo,
        IFormFile? footerLogo)
    {
        byte[]? headerLogoBytes = null;
        if (headerLogo != null && headerLogo.Length > 0)
        {
            using var ms = new MemoryStream();
            await headerLogo.CopyToAsync(ms);
            headerLogoBytes = ms.ToArray();
        }

        byte[]? footerLogoBytes = null;
        if (footerLogo != null && footerLogo.Length > 0)
        {
            using var ms = new MemoryStream();
            await footerLogo.CopyToAsync(ms);
            footerLogoBytes = ms.ToArray();
        }

        string sanitizedTitle = string.IsNullOrWhiteSpace(model.ReportTitle)
            ? "Report"
            : string.Concat(model.ReportTitle.Split(Path.GetInvalidFileNameChars()));

        if (model.Format == OutputFormat.Excel)
        {
            byte[] excelBytes = _reportService.GenerateExcel(model);
            return File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{sanitizedTitle}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }
        else
        {
            byte[] pdfBytes = _reportService.GeneratePdf(model, headerLogoBytes, footerLogoBytes);
            return File(
                pdfBytes,
                "application/pdf",
                $"{sanitizedTitle}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");
        }
    }

    [HttpGet]
    public IActionResult GetFonts()
    {
        return Json(AvailableFonts);
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}

public sealed class RandomDataRequest
{
    public int Count { get; set; } = 10;
}

public sealed class ErrorViewModel
{
    public string? RequestId { get; set; }
    public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
}

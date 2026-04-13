using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace RaporlamaPortali.Pages;

[AllowAnonymous]
public class GirisModel : PageModel
{
    public bool Hata { get; private set; }
    public string ReturnUrl { get; private set; } = "/";

    public void OnGet()
    {
        Hata = Request.Query["hata"].FirstOrDefault() == "1";
        ReturnUrl = Request.Query["ReturnUrl"].FirstOrDefault() ?? "/";
        if (!ReturnUrl.StartsWith("/")) ReturnUrl = "/";
    }
}

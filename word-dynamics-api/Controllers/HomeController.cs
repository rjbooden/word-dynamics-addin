using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace word_dynamics_api.Controllers;

public class HomeController : ControllerBase
{

    // Just having some fun, and need some default
    static string[] Smileys = new string [] { "๐","๐","๐","๐คฃ","๐","๐","๐","๐","๐",
        "๐","๐","๐","๐ฅฒ","๐","๐ค","๐ค","๐ซก","๐คจ","๐","๐","๐","๐","๐ฃ","๐ฅ","๐ฎ",
        "๐ค","๐ฏ","๐ช","๐ซ","๐ฅฑ","๐ด","๐","๐","๐","๐","๐คค","๐","๐","๐","๐","๐",
        "๐ซ ","๐ฒ","โน๏ธ","๐","๐","๐","๐","๐ค","๐ข","๐ญ","๐ฆ","๐ง","๐จ","๐ฉ","๐คฏ","๐ฌ",
        "๐ฎโ๐จ","๐ฐ","๐ฑ","๐ฅต","๐ฅถ","๐ณ","๐คช","๐ต","๐ตโ๐ซ","๐ฅด","๐ ","๐ก","๐คฌ","๐ท","๐ค","๐ค",
        "๐คข","๐คฎ","๐คง","๐"};

    [AllowAnonymous]
    [HttpGet]
    public IActionResult Index()
    {
        Random rnd = new Random();
        int newSmiley = rnd.Next(0, Smileys.Length - 1);
        return Content(Smileys[newSmiley]);
    }
}
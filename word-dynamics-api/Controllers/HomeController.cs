using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace word_dynamics_api.Controllers;

public class HomeController : ControllerBase
{

    // Just having some fun, and need some default
    static string[] Smileys = new string [] { "😀","😁","😂","🤣","😃","😄","😅","😆","😉",
        "😊","😋","😎","🥲","🙂","🤗","🤔","🫡","🤨","😐","😑","🙄","😏","😣","😥","😮",
        "🤐","😯","😪","😫","🥱","😴","😌","😛","😜","😝","🤤","😒","😓","😔","😕","🙃",
        "🫠","😲","☹️","🙁","😖","😞","😟","😤","😢","😭","😦","😧","😨","😩","🤯","😬",
        "😮‍💨","😰","😱","🥵","🥶","😳","🤪","😵","😵‍💫","🥴","😠","😡","🤬","😷","🤒","🤕",
        "🤢","🤮","🤧","😇"};

    [AllowAnonymous]
    [HttpGet]
    public IActionResult Index()
    {
        Random rnd = new Random();
        int newSmiley = rnd.Next(0, Smileys.Length - 1);
        return Content(Smileys[newSmiley]);
    }
}
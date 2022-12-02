using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using FlaUI.UIA2;
using FlaUI.Core.Tools;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;

string path = args[0];
string extension = string.Concat("*", args[1]);
string dir = Path.GetFileName(path);
List<string> fileswithpath = new List<string>();
List<string> files = new List<string>();
fileswithpath.AddRange(Directory.GetFiles(path, extension));

foreach(string file in fileswithpath)
{
    string f = Path.GetFileName(file);
    files.Add(f);
    Console.WriteLine(f);
}

var app = FlaUI.Core.Application.Launch("explorer.exe", path);

try 
{ 
    using (var automation = new UIA2Automation())
    {
        var desktop = automation.GetDesktop();
        var window = Retry.WhileNull<Window>(() =>
        {
            return desktop.FindFirstChild(cf => cf.ByName(dir)).AsWindow();
        }, timeout: TimeSpan.FromSeconds(10), throwOnTimeout: true, timeoutMessage: "can't get the window");
        if (window.Success)
        {
            var list = Retry.WhileNull<ListBox>(() =>
            {
                //return window.Result.FindFirstByXPath($"/Pane[3]/Pane/Pane[2]/List").AsListBox();
                return window.Result.FindFirstDescendant(cf => cf.ByName("Items View")).AsListBox();
            }, TimeSpan.FromMinutes(2), throwOnTimeout: true);
            if (list.Success)
            {

                foreach (var litem in list.Result.Items)//list.Result.FindAllByXPath("/ListItem"))
                {
                    Console.WriteLine(litem.Name);
                    if (files.Contains(litem.Name))
                    {
                        litem.AddToSelection();
                    }
                }
                list.Result.SelectedItem.RightClick();

                var menu = Retry.WhileNull<Menu>(() =>
                {
                    return desktop.FindFirstChild(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Menu)).AsMenu();
                }, TimeSpan.FromSeconds(50), throwOnTimeout: true);
                if (menu.Success)
                {
                    var nasca = menu.Result.Items.First(it => it.Name == "NASCA +SD");

                    nasca.Focus();
                    nasca.Expand();

                    var denasca = Retry.WhileNull<MenuItem>(() =>
                    {
                        return nasca.FindFirstDescendant(cf => cf.ByName("Convert to decrypted document")).AsMenuItem();
                    }, TimeSpan.FromSeconds(50), throwOnTimeout:true);
                    if(denasca.Success)
                    {
                        Console.WriteLine(denasca.Result.Name);
                        denasca.Result.Click();
                    }
                }
                
            }
            window.Result.Close();

        } 
    }
    app.Close();
}
catch(Exception e)
{
    app.Close();
}
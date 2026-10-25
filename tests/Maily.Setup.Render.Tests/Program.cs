using Maily.OutlookAddin.Setup;
using System.Drawing;
using System.Windows.Forms;
internal static class RenderProgram
{
    [STAThread] static void Main(string[] args)
    {
        if (args.Length != 1) throw new ArgumentException("Supply a unique output folder for retained screenshots.");
        Directory.CreateDirectory(args[0]);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        var options = SetupOptions.Parse([]);
        foreach (var scenario in new[] { "welcome", "connect", "restart", "support", "remove", "ready" })
        {
            using var form = new SetupForm(options with { Uninstall = scenario == "remove" }, "Local", null);
            switch (scenario) {
                case "connect": form.ApplyInstallationResult(20,true); break;
                case "restart": form.ApplyInstallationResult(21,true); break;
                case "support": form.ApplyInstallationResult(23,false); break;
                case "remove": form.ApplyRemovalResult(22); break;
                case "ready": form.ApplyInstallationResult(0,true); break;
            }
            // Show only this synthetic form off-screen so WinForms creates and paints child handles.
            // It is never interactive and does not invoke any setup actions.
            form.ShowInTaskbar = false;
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new Point(-10000,-10000);
            foreach (var button in form.Controls.OfType<Button>()) button.Enabled = false;
            form.Show(); Application.DoEvents(); form.PerformLayout();
            var expected = scenario switch { "connect" => "Connect Outlook", "restart" => "Check connection", "support" => "Close", "remove" => "Finish removal", "ready" => "Open Outlook", _ => "Install" };
            if (!form.Controls.OfType<Button>().Any(button => button.Text == expected)) throw new Exception(scenario+" has the wrong next action.");
            if (form.Controls.OfType<TextBox>().Any(box => box.Visible)) throw new Exception("Technical logs must start collapsed.");
            if (scenario != "remove" && form.Controls.OfType<ListView>().Single().Items.Count != 3) throw new Exception("Setup must present three stages.");
            if (form.ClientSize.Height > 600) throw new Exception("Default setup does not fit a common laptop viewport.");
            using var bitmap = new Bitmap(form.Width,form.Height);
            form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));
            var paintedPixels = 0;
            for(var y=80;y<Math.Min(500,bitmap.Height);y++) for(var x=25;x<Math.Min(740,bitmap.Width);x++)
                if(bitmap.GetPixel(x,y).GetBrightness()<0.7f) paintedPixels++;
            if(paintedPixels<1000) throw new Exception("Rejected blank/unpainted client area: "+scenario);
            bitmap.Save(Path.Combine(args[0],scenario+".png"));
            Console.WriteLine("PASS real form render: "+scenario);
        }

        foreach(var repair in new[]{false,true})
        {
            using var owner = new SetupForm(options,"Local",null);
            owner.ApplyInstallationResult(20, !repair, repair ? "repair-required" : "activation-required");
            using var guide = owner.CreateConnectionGuide();
            guide.ShowInTaskbar=false; guide.StartPosition=FormStartPosition.Manual; guide.Location=new Point(-10000,-10000);
            foreach(var button in guide.Controls.OfType<Button>()) button.Enabled=false;
            guide.Show(); Application.DoEvents(); guide.PerformLayout();
            if(!guide.Controls.OfType<Button>().Any(button=>button.Text==(repair?"Continue setup":"Check connection"))) throw new Exception("Guide routes to the wrong phase");
            using var image = new Bitmap(guide.Width,guide.Height); guide.DrawToBitmap(image,new Rectangle(Point.Empty,image.Size));
            image.Save(Path.Combine(args[0],repair?"repair-guide.png":"connection-guide.png"));
            Console.WriteLine("PASS real form render: "+(repair?"repair guide":"connection guide"));
        }
        Console.WriteLine("RESULT: 8 real-form scenarios passed; no host discovery, Outlook, account, registry or install actions.");
    }
}

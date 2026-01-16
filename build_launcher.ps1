# build_launcher.ps1
# Builds COSA.exe which runs cosa.ps1 in the same folder.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$src = @'
using System;
using System.Diagnostics;
using System.IO;

class Program
{
    static int Main(string[] args)
    {
        try
        {
            var root = AppDomain.CurrentDomain.BaseDirectory;
            var ps1 = Path.Combine(root, "cosa.ps1");
            if (!File.Exists(ps1))
            {
                Console.Error.WriteLine("Missing cosa.ps1 next to the EXE.");
                return 2;
            }

            var psi = new ProcessStartInfo();
            psi.FileName = "powershell.exe";
            psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + ps1 + "\"";
            psi.UseShellExecute = false;
            psi.CreateNoWindow = false; // show console window
            psi.WorkingDirectory = root;

            using (var p = Process.Start(psi))
            {
                p.WaitForExit();
                return p.ExitCode;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
    }
}
'@

$outExe = Join-Path $root "COSA.exe"
Add-Type -TypeDefinition $src -Language CSharp -OutputAssembly $outExe -OutputType ConsoleApplication
Write-Host "Built: $outExe"

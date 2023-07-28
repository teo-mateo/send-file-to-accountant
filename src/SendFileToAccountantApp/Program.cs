// See https://aka.ms/new-console-template for more information

using System.ComponentModel;
using System.Diagnostics;
using System.Security.Principal;
using Heapzilla.Common.Filesystem;
using MailKit.Net.Smtp;
using Microsoft.Extensions.Configuration;
using MimeKit;

// current executable name: sfta.exe    

var appSettingsJsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json").ThrowIfFileNotExists();

var configuration = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile(appSettingsJsonPath, optional: false, reloadOnChange: true)
    .Build();

switch (args)
{
    case []: 
        Console.WriteLine("No arguments specified. \r\nUse --install to install the shell integration or --uninstall to uninstall it. \r\nUse --send <file> to send a file.");
        return;
    case ["--install"] or ["-i"]:
        Install();
        return;
    case ["--uninstall"] or ["-u"]:
        Uninstall();
        return;
    case ["--send", _] or ["-s", _]:
        Send();
        return;
}

// Install registry keys
void Install()
{
    try
    {
        // Path to your .reg file
        var pathToRegFile = AppDomain.CurrentDomain.BaseDirectory + "shell-integration.reg";

        // replace {{applicationPath}} in the shellintegration.reg file with the path to the current executable
        // it must be between double quotes and the backslashes should be escaped
        var applicationPath = AppDomain.CurrentDomain.BaseDirectory + "sfta.exe";
        var regText = File.ReadAllText(pathToRegFile);
        if (regText.Contains("{{applicationPath}}"))
        {
            
            regText = regText.Replace("{{applicationPath}}", "\\\"" + applicationPath.Replace("\\", "\\\\")+ "\\\"");
            File.WriteAllText(pathToRegFile, regText);            
        }
        
        var process = new Process();
        var startInfo = new ProcessStartInfo
        {
            WindowStyle = ProcessWindowStyle.Hidden,
            FileName = "cmd.exe",
            Arguments = "/C regedit.exe /s " + pathToRegFile
        };
        process.StartInfo = startInfo;
        process.Start();
        
        // wait for process to end and if it was not successful throw an exception
        process.WaitForExit();
        if (process.ExitCode != 0)
            throw new Win32Exception(process.ExitCode, "An error occurred while installing the registry keys. Exit code: " + process.ExitCode.ToString());
    }
    catch (Exception ex)
    {
        Console.WriteLine("An error occurred: " + ex.Message);
    }
}

void Uninstall()
{
    var isAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent())
        .IsInRole(WindowsBuiltInRole.Administrator);

    if (!isAdmin)
    {
        Console.WriteLine("You must run this application as administrator to remove the shell integration.");
        return;
    }
    
    // start "uninst.bat" with an elevated command prompt
    var process = new Process();
    var startInfo = new ProcessStartInfo
    {
        WindowStyle = ProcessWindowStyle.Hidden,
        FileName = "cmd.exe",
        Arguments = "/C uninst.bat"
    };
    process.StartInfo = startInfo;
    process.Start();
    
    // wait for process to end and if it was not successful throw an exception
    process.WaitForExit();
    if (process.ExitCode != 0)
        throw new Win32Exception(process.ExitCode, "An error occurred while uninstalling the registry keys. Exit code: " + process.ExitCode.ToString());
    Console.WriteLine("Registry keys successfully uninstalled.");
}

void Send()
{
    // if the first argument is not a file path or the file does not exist, throw an exception
    var attachment = args[1].ThrowIfFileNotExists();
    var server = configuration["Server"] ?? throw new Exception("Server not specified");
    var port = configuration["Port"] ?? throw new Exception("Port not specified");
    var fromName = configuration["FromName"] ?? throw new Exception("FromName not specified");
    var fromAddress = configuration["FromAddress"] ?? throw new Exception("FromAddress not specified");
    var toName = configuration["ToName"] ?? throw new Exception("ToName not specified");
    var toAddress = configuration["ToAddress"] ?? throw new Exception("ToAddress not specified");
    var username = configuration["Username"] ?? throw new Exception("Username not specified");
    var password = configuration["Password"] ?? throw new Exception("Password not specified");

    var subject = configuration["Subject"] ?? throw new Exception("Subject not specified");
    
    var message = new MimeMessage();
    message.From.Add(new MailboxAddress(fromName, fromAddress));
    message.To.Add(new MailboxAddress(toName, toAddress));
    message.Bcc.Add(new MailboxAddress(fromName, fromAddress));
    
    message.Subject = subject.Replace("{{filename}}", Path.GetFileName(attachment));

    var bodyBuilder = new BodyBuilder
    {
        HtmlBody = "Heapzilla BV"
    };
    bodyBuilder.Attachments.Add(attachment);
    message.Body = bodyBuilder.ToMessageBody();

    using var client = new SmtpClient();
    client.Connect(server, int.Parse(port), true);
    client.Authenticate(username, password);
    client.Send(message);
    client.Disconnect(true);
    
    // rename the attachment file by prefixing the filename with (sent YYYY-mm-dd), unless already prefixed
    var newFilename = $"(sent {DateTime.Now:yyyy-MM-dd}) {Path.GetFileName(attachment)}";
    if (!Path.GetFileName(attachment).StartsWith("(sent"))
        File.Move(attachment, Path.Combine(Path.GetDirectoryName(attachment)!, newFilename));
    

}
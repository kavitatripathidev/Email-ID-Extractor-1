/*
tool that will go through all emails in a mailbox (gmail) and dump 
all email addresses regardless of whether those addresses are saved as contacts.
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using System.Threading;
using System.IO;
using MailKit.Search;

namespace EmailIDsExtractor
{
    public class EmailAddress
    {
        public string Name { get; set; }
        public string Address { get; set; }
    }
    public class EmailConfiguration
    {
        public string ImapServer { get; set; }
        public int ImapPort { get; set; }
        public string ImapUsername { get; set; }
        public string ImapPassword { get; set; }
        public EmailConfiguration()
        {
            try
            {
                XElement conn = XElement.Load("connection.config");
                ImapServer = (string)conn.Element("Server");
                ImapPort = (int)conn.Element("Port");
                ImapUsername = (string)conn.Element("UserName");
                ImapPassword = (string)conn.Element("Password");
                if (ImapServer == null || ImapPort == 0)
                    throw new IOException("Connection file could not be found or have some empty values.");
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
            }
        }
    }
    public class EmailIDsExtractor
    {
        private EmailConfiguration emailConfiguration = new EmailConfiguration();
        private Dictionary<string, EmailAddress> emailIDs;
        public string outputFolder = "output";
        public string fileName = "Email IDs";
        public string fileExtension = ".csv";
        public string tempFilePath = "temp.tmp";
        private int lastEmailRead = 0;
        private int filePartNo = 0;
        private int tempCount = 0;
        private void SetTempFileData(int lastEmailRead, int filePartNo, int tempCount)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(tempFilePath))
                {
                    sw.WriteLine(lastEmailRead);
                    sw.WriteLine(filePartNo);
                    this.tempCount += tempCount;
                    sw.WriteLine(this.tempCount);
                }
            }
            catch (IOException ex)
            {
                throw ex;
            }
        }
        private void RestoreTempFileData()
        {
            try
            {
                if (File.Exists(tempFilePath))
                    using (StreamReader sr = new StreamReader(tempFilePath))
                    {
                        lastEmailRead = int.Parse(sr.ReadLine());
                        filePartNo = int.Parse(sr.ReadLine());
                        tempCount = int.Parse(sr.ReadLine());
                        DataExistedInTemp = true;
                    }
                
            }
            catch
            {
                //
            }
        }
        private void CheckTempOutputFile()
        {
            System.Console.WriteLine("Output will be overwritten if files exists in output folder.");
            System.Console.WriteLine("\nReading addresses...");
            if (!DataExistedInTemp && File.Exists(Path.Combine(outputFolder, "temp output") + fileExtension))
                File.Delete(Path.Combine(outputFolder, "temp output") + fileExtension);
        }
        public List<EmailAddress> GetEmailAddresses()
        {
            try
            {
                using (var client = new ImapClient())
                {
                    client.Connect(emailConfiguration.ImapServer, emailConfiguration.ImapPort, SecureSocketOptions.SslOnConnect);

                    client.Authenticate(emailConfiguration.ImapUsername, emailConfiguration.ImapPassword);

                    client.Inbox.Open(FolderAccess.ReadOnly);
                    emailIDs = new Dictionary<string, EmailAddress>();
                    List<EmailAddress> tempEmailIDs = new List<EmailAddress>();

                    var inbox = client.Inbox;
                    System.Console.WriteLine("Fetching Emails...");
                    System.Console.WriteLine(inbox.Count + " emails found in Inbox.\nTemporary output files can be found in output\\temp folder.");
                    var uids = client.Inbox.Search(SearchQuery.All);

                    RestoreTempFileData();
                    CheckTempOutputFile();

                    for (int i = lastEmailRead; i < uids.Count; i++)
                    {
                        var headers = inbox.GetHeaders(uids[i]);
                        var addresses = headers["From"] + "," + headers["To"] + "," + headers["CC"];
                        var allAddresses = addresses.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                        foreach (var address in allAddresses)
                        {
                            var add = address.Split('<');
                            string addr = (add.Count() > 1) ? add[1].TrimEnd('>') : add[0];
                            string name = (add.Count() > 1) ? add[0] : "";
                            if (!emailIDs.ContainsKey(addr))
                            {
                                var email = new EmailAddress();
                                email.Address = addr;
                                email.Name = name;
                                tempEmailIDs.Add(email);
                                emailIDs.Add(addr, email);
                            }
                        }
                        if (tempEmailIDs.Count > 100)
                        {
                            tempEmailIDs.RemoveAll(x => x.Address == emailConfiguration.ImapUsername);
                            WriteToFile(tempEmailIDs, "temp output");
                            WriteToFile(tempEmailIDs, filePartNo);
                            System.Console.WriteLine($"Extracted {tempEmailIDs.Count} emails.");
							filePartNo++;
                            SetTempFileData((i+1), filePartNo, tempEmailIDs.Count);                            
                            tempEmailIDs.Clear();
                        }
                    }
                    client.Disconnect(true);
                    if (tempEmailIDs.Any())
                    {
                        WriteToFile(tempEmailIDs, "temp output");
                        WriteToFile(tempEmailIDs, filePartNo);
                        System.Console.WriteLine("All emails read from Inbox.");
                    }
                    RestoreExtractedEmails();
                    return emailIDs.Values.ToList();
                }
            }
            catch (AuthenticationException ex)
            {
                System.Console.WriteLine(ex.Message);
                System.Console.WriteLine("Please check the configuration details in the config file.");
            }
            catch (System.Net.Sockets.SocketException ex)
            {
                System.Console.WriteLine("Could not connect to the email server!");
                System.Console.WriteLine("Possible reasons:\n Network issue.\n Incorrect configuration details.");
                System.Console.WriteLine(ex.Message);
            }
            return null;

        }

        public void WriteToFile(List<EmailAddress> emailAddresses, int fileNo)
        {
            try
            {
                string tempPath = Path.Combine(outputFolder, "temp");
                if (!Directory.Exists(tempPath))
                    Directory.CreateDirectory(tempPath);
                using (StreamWriter sw = new StreamWriter(Path.Combine(tempPath, fileName) + fileNo + fileExtension))
                {
                    foreach (var email in emailAddresses)
                    {
                        sw.WriteLine(email.Name + "," + email.Address);
                    }
                }
            }
            catch (IOException ex)
            {
                System.Console.WriteLine(ex.Message);
            }
        }
        public void WriteToFile(List<EmailAddress> emailAddresses, string fileName)
        {
            try
            {
                string tempOutputPath = Path.Combine(outputFolder, fileName) + fileExtension;
                
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);
                using (StreamWriter sw = new StreamWriter(tempOutputPath, true))
                {
                    foreach (var email in emailAddresses)
                    {
                        sw.WriteLine(email.Name + "," + email.Address);
                    }
                }
            }
            catch (IOException ex)
            {
                System.Console.WriteLine(ex.Message);
            }
        }

        private bool DataExistedInTemp = false;
        void RestoreExtractedEmails()
        {
            try
            {
                if (DataExistedInTemp)
                {
                    System.Console.WriteLine("Temporary output file was found. Merging previously read data...");
                    if (File.Exists(Path.Combine(outputFolder, "temp output") + fileExtension))
                        using (StreamReader sr = new StreamReader(Path.Combine(outputFolder, "temp output") + fileExtension))
                        {
                            int i = 0;
                            string line;
                            while ((line = sr.ReadLine()) != null)
                            {
                                var email = new EmailAddress();
                                email.Name = line.Split(',')[0];
                                email.Address = line.Split(',')[1];
                                emailIDs.Add(email.Address, email);
                                i++;
                                if (i > tempCount)
                                    break;
                            }
                        }
                }

            }
            catch
            {
                //
            }
        }
    }
    class Program
    {
        static EmailIDsExtractor extractor = new EmailIDsExtractor();
        static void Main(string[] args)
        {
            System.Console.WriteLine("Program started.");
            try
            {
                var emailAddresses = extractor.GetEmailAddresses();
                if (emailAddresses.Any())
                {
                    extractor.WriteToFile(emailAddresses, extractor.fileName);
                    System.Console.WriteLine($"\n Output File: {Path.GetFullPath(Path.Combine(extractor.outputFolder, extractor.fileName) + extractor.fileExtension)}");
					File.Delete(Path.Combine(extractor.outputFolder + "temp output" + extractor.fileExtension));
                    File.Delete(extractor.tempFilePath);
                }
                else
                {
                    System.Console.WriteLine("No emails found!");
                }
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
            }
            System.Console.WriteLine("\nPress any key to close.");
            Console.ReadKey();
        }
    }
}

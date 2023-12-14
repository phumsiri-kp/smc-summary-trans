using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Mail;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using ClosedXML.Excel;
using System.IO;
using System.Threading;

namespace smcSummaryTrans
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string mailServer = "142.1.10.30";
            string fromMail = "no-reply@kingpower.com";
            string toMail = ConfigurationManager.AppSettings["toEmail"];
            string subject = "[Auto mail][" + DateTime.Now.ToString("dd/MM/yyyy") + "] Daily summary data for reconcile SMC and Mulesoft";
            string bodyMail = "<h3>Daily summary data for reconcile SMC and Mulesoft</h3><br>";
            bool isBodyHtml = true;
            List<string> attachFile = new List<string>();

            string sqlQuery;
            DataTable resultDataTable;
            SqlQueryFunction sqlQueryExecutor = new SqlQueryFunction();

            // SMC Account Sync
            bodyMail += "<p>SMC Account Sync</p>";
            sqlQuery = "SELECT [Date] ,[Total] ,[Success] ,[Fail] FROM [Newmember].[dbo].[v_member_sum_triggerToMulesoft_all];";
            resultDataTable = sqlQueryExecutor.ExecuteQueryWithRetry(sqlQuery);
            bodyMail += ConvertDataTableToHtml(resultDataTable);

            // New Registration
            bodyMail += "<br><p>New Registration</p>";
            sqlQuery = "SELECT [Date] ,[Total] ,[Success] ,[Fail] FROM [Newmember].[dbo].[v_member_sum_triggerToMulesoft_insert];";
            resultDataTable = sqlQueryExecutor.ExecuteQueryWithRetry(sqlQuery);
            bodyMail += ConvertDataTableToHtml(resultDataTable);

            // lv and spend
            bodyMail += "<br><p>LV transaction</p>";
            sqlQuery = "SELECT [Date],[Total] FROM [Newmember].[dbo].[v_member_sum_lvSpendTrans_all];";
            resultDataTable = sqlQueryExecutor.ExecuteQueryWithRetry(sqlQuery);
            bodyMail += ConvertDataTableToHtml(resultDataTable);

            // Co-brand
            bodyMail += "<br><p>Co-Brand</p>";
            sqlQuery = "SELECT [Date],[Total],[KBankClose] FROM [Newmember].[dbo].[v_member_sum_cobrandTrans_all];";
            resultDataTable = sqlQueryExecutor.ExecuteQueryWithRetry(sqlQuery);
            bodyMail += ConvertDataTableToHtml(resultDataTable);

            // Error data
            sqlQuery = "SELECT [transType],[member_id],[shopping_card],[response_time],[response_message] FROM [Newmember].[dbo].[v_member_error_trigger_trans];";
            DataTable errorResult = sqlQueryExecutor.ExecuteQueryWithRetry(sqlQuery);

            // Convert DataTable to XLSX
            byte[] excelFileBytes = DataTableToExcel(errorResult);

            // Get the current directory path
            string currentDirectory = Environment.CurrentDirectory;
            //string excelFilePath = "SMC_MemberProfile_Trigger_Error_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            string excelFilePath = "SMC_MemberProfile_Trigger_Error.xlsx";
            string fullPath = System.IO.Path.Combine(currentDirectory, excelFilePath);

            // Save XLSX content to a file
            if (excelFileBytes != null) 
            { 
                File.WriteAllBytes(fullPath, excelFileBytes);
                attachFile.Add(fullPath);
            }

            // send email
            var rs = SendMail(mailServer, fromMail, toMail, subject, bodyMail, isBodyHtml, attachFile);
            Console.WriteLine(rs);

            // Delete the XLSX file after sending the email
            //DeleteXLSXFile(fullPath);
        }
        static string ConvertDataTableToHtml(DataTable dataTable)
        {
            StringBuilder htmlStringBuilder = new StringBuilder();

            // Start HTML table
            htmlStringBuilder.AppendLine("<table border='1'>");

            // Add table header
            htmlStringBuilder.AppendLine("<tr>");
            foreach (DataColumn column in dataTable.Columns)
            {
                htmlStringBuilder.AppendLine($"<th>{column.ColumnName}</th>");
            }
            htmlStringBuilder.AppendLine("</tr>");

            // Add table rows
            foreach (DataRow row in dataTable.Rows)
            {
                htmlStringBuilder.AppendLine("<tr>");
                foreach (var item in row.ItemArray)
                {
                    htmlStringBuilder.AppendLine($"<td>{item}</td>");
                }
                htmlStringBuilder.AppendLine("</tr>");
            }

            // End HTML table
            htmlStringBuilder.AppendLine("</table>");

            return htmlStringBuilder.ToString();
        }

        static byte[] DataTableToExcel(DataTable dataTable)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // Write header
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cell(1, col + 1).Value = dataTable.Columns[col].ColumnName;
                }

                // Write data
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cell(row + 2, col + 1).Value = dataTable.Rows[row][col].ToString();
                    }
                }

                // Save the workbook to a memory stream
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        static string QuoteField(string field)
        {
            // If the field contains a comma, quote it
            if (field.Contains(","))
                return $"\"{field}\"";
            else
                return field;
        }

        static void DeleteXLSXFile(string filePath)
        {
            int maxAttempts = 3;
            int attemptDelayMilliseconds = 1000;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    // Check if the file exists before attempting to delete
                    if (File.Exists(filePath))
                    {
                        // Ensure the file is not open or in use
                        using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                        {
                            // Add your logic here if you need to perform operations on the file before deletion
                        }

                        // Attempt to delete the file after ensuring it's not in use
                        File.Delete(filePath);
                        Console.WriteLine("File deleted successfully.");
                        break; // Exit the loop if deletion is successful
                    }
                    else
                    {
                        Console.WriteLine("File does not exist.");
                        break; // Exit the loop if the file does not exist
                    }
                }
                catch (IOException ex)
                {
                    Console.WriteLine($"Error deleting the file (Attempt {attempt}/{maxAttempts}): {ex.Message}");

                    // Sleep before the next attempt
                    Thread.Sleep(attemptDelayMilliseconds);
                }
            }
        }

        public static string SendMail(string mailServer, string fromMail, string toMail, string subject, string bodyMail, bool isBodyHtml = false, List<string> attachFile = null)
        {
            string result = string.Empty;
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(mailServer);//, 587);//("smtp.gmail.com");
                mail.From = new MailAddress(fromMail);
                //mail.To.Add(toMail);
                foreach (var recipient in toMail.Split(','))
                {
                    mail.To.Add(recipient.Trim());
                }
                mail.Subject = subject;
                mail.IsBodyHtml = isBodyHtml;
                mail.Body = bodyMail;

                if (attachFile != null && attachFile.Count > 0)
                {
                    foreach (string att in attachFile)
                    {
                        System.Net.Mail.Attachment attachment;
                        attachment = new System.Net.Mail.Attachment(att);//("c:/textfile.txt");
                        mail.Attachments.Add(attachment);
                    }

                }
                SmtpServer.Send(mail);
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }

            return result;
        }

    }

    public class SqlQueryFunction
    {
        private string connectionString = "Data Source=member-server;Initial Catalog=Newmember;User ID=sa;Password=sql2000;";

        public DataTable ExecuteQueryWithRetry(string sqlCommandText)
        {
            int maxRetries = int.Parse(ConfigurationManager.AppSettings["max_retry"].ToString());
            int retryCount = 0;

            while (retryCount < maxRetries)
            {
                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();

                        using (SqlCommand command = new SqlCommand(sqlCommandText, connection))
                        {

                            // Create a DataTable to hold the results
                            DataTable resultTable = new DataTable();

                            // Use a DataAdapter to fill the DataTable
                            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                            {
                                adapter.Fill(resultTable);
                            }

                            return resultTable; // Query successful, return the DataTable
                        }
                    }
                }
                catch (SqlException ex)
                {
                    // Handle the exception, you might want to log it or perform some specific actions
                    Console.WriteLine($"Error: {ex.Message}");

                    // Increment the retry count
                    retryCount++;

                    // Wait for a short period before the next retry (optional)
                    Thread.Sleep(1000);
                }
            }

            return null; // Unable to execute the query after maximum retries
        }
    }
}

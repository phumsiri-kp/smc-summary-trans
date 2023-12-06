using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Mail;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace smcSummaryTrans
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string mailServer = "142.1.10.30";
            string fromMail = "no-reply@kingpower.com";
            string toMail = ConfigurationManager.AppSettings["toEmail"];
            string subject = "[Auto mail]["+ DateTime.Now.ToString("dd/MM/yyyy") +"] Daily summary data for reconcile SMC and Mulesoft"; 
            string bodyMail = "<h3>Daily summary data for reconcile SMC and Mulesoft</h3><br>";
            bool isBodyHtml = true;
            List<string> pdf = new List<string>();

            string sqlQuery;
            DataTable resultDataTable;
            SqlQueryFunction sqlQueryExecutor = new SqlQueryFunction();

            // SMC Account Sync
            bodyMail += "<p>SMC Account Sync</p>";
            sqlQuery = "SELECT [Date] ,[Total] ,[Success] ,[Fail] FROM [Newmember].[dbo].[v_member_sum_triggerToMulesoft_all];";
            resultDataTable = sqlQueryExecutor.QueryDataWithTransaction(sqlQuery);
            bodyMail += ConvertDataTableToHtml(resultDataTable);

            // New Registration
            bodyMail += "<br><p>New Registration</p>";
            sqlQuery = "SELECT [Date] ,[Total] ,[Success] ,[Fail] FROM [Newmember].[dbo].[v_member_sum_triggerToMulesoft_insert];";
            resultDataTable = sqlQueryExecutor.QueryDataWithTransaction(sqlQuery);
            bodyMail += ConvertDataTableToHtml(resultDataTable);

            // lv and spend
            bodyMail += "<br><p>LV transaction</p>";
            sqlQuery = "SELECT [Date],[Total] FROM [Newmember].[dbo].[v_member_sum_lvSpendTrans_all];";
            resultDataTable = sqlQueryExecutor.QueryDataWithTransaction(sqlQuery);
            bodyMail += ConvertDataTableToHtml(resultDataTable);

            // Co-brand
            bodyMail += "<br><p>Co-Brand</p>";
            sqlQuery = "SELECT [Date],[Total],[KBankClose] FROM [Newmember].[dbo].[v_member_sum_cobrandTrans_all];";
            resultDataTable = sqlQueryExecutor.QueryDataWithTransaction(sqlQuery);
            bodyMail += ConvertDataTableToHtml(resultDataTable);

            // send email
            var rs = SendMail(mailServer, fromMail, toMail, subject, bodyMail, isBodyHtml, pdf);
            Console.WriteLine(rs);
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

        public DataTable QueryDataWithTransaction(string sqlQuery)
        {
            DataTable resultDataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Start a transaction
                SqlTransaction transaction = connection.BeginTransaction();

                try
                {
                    // Execute your SQL query within the transaction
                    using (SqlCommand command = connection.CreateCommand())
                    {
                        command.Transaction = transaction;
                        command.CommandText = sqlQuery;

                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            // Fill the DataTable with the result of the query
                            adapter.Fill(resultDataTable);
                        }
                    }

                    // Commit the transaction if everything is successful
                    transaction.Commit();
                    Console.WriteLine("Transaction committed successfully!");
                }
                catch (Exception ex)
                {
                    // Handle exceptions and roll back the transaction in case of an error
                    Console.WriteLine($"Error: {ex.Message}");
                    transaction.Rollback();
                }
                finally
                {
                    // Close the connection in the finally block to ensure it's always closed
                    connection.Close();
                    Console.WriteLine("Connection closed.");
                }
            }

            return resultDataTable;
        }
    }
}

using Quartz;
using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace WindowsServiceCS
{
	public class ADJob2 : IJob
	{
		public ADJob2()
		{
		}

		public void Execute(IJobExecutionContext context)
		{
			try
			{
				string constrj1 = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
				SqlConnection mainConnectionj1 = new SqlConnection(constrj1);
				mainConnectionj1.Open();
				string EmailAddress = string.Empty;
				string Password = string.Empty;
				string SmtpServer = string.Empty;
				int Port = 0;
				SqlConnection conj2 = new SqlConnection(constrj1);
				conj2.Open();
				SqlCommand cmdj2 = new SqlCommand("select * from dbo.AbpSettings", conj2);
				DataTable dtj2 = new DataTable();
				using (SqlDataAdapter sdaj1 = new SqlDataAdapter(cmdj2))
				{
					sdaj1.Fill(dtj2);
				}
				foreach (DataRow datarow2 in dtj2.Rows)
				{
					if (datarow2["Name"].ToString() == "Abp.Net.Mail.Smtp.Port")
					{
						Port = Convert.ToInt32(datarow2["Value"]);
					}
					else if (datarow2["Name"].ToString() == "Abp.Net.Mail.Smtp.UserName")
					{
						EmailAddress = datarow2["Value"].ToString();
					}
					else if (datarow2["Name"].ToString() == "Abp.Net.Mail.Smtp.Password")
					{
						Password = datarow2["Value"].ToString();
					}
					else if (datarow2["Name"].ToString() == "Abp.Net.Mail.Smtp.Host")
					{
						SmtpServer = datarow2["Value"].ToString();
					}
				}
				bool flag = false;
				string queryj1 = string.Concat("select * from dbo.EmailDetail where IsSent='", flag.ToString(), "' and EmailStatus is null");
				DataTable dtj1 = new DataTable();
				using (SqlConnection conj1 = new SqlConnection(constrj1))
				{
					using (SqlCommand cmdj1 = new SqlCommand(queryj1))
					{
						cmdj1.Connection = conj1;
						using (SqlDataAdapter sda = new SqlDataAdapter(cmdj1))
						{
							sda.Fill(dtj1);
						}
					}
				}
				foreach (DataRow datarow1 in dtj1.Rows)
				{
					string From = string.Empty;
					string To = string.Empty;
					string Cc = string.Empty;
					string Subject = string.Empty;
					string MailBody = string.Empty;
					if (datarow1["ToEmailAddress"].ToString() != null)
					{
						To = datarow1["ToEmailAddress"].ToString();
					}
					if (datarow1["CcEmailAddress"].ToString() != null)
					{
						Cc = datarow1["CcEmailAddress"].ToString();
					}
					if (datarow1["Subject"].ToString() != null)
					{
						Subject = datarow1["Subject"].ToString();
					}
					if (datarow1["MailBody"].ToString() != null)
					{
						MailBody = datarow1["MailBody"].ToString();
					}
					if (Port == 0)
					{
						Port = 25;
					}
					try
					{
						this.WriteToFile("Error 1");
						using (MailMessage mm = new MailMessage(EmailAddress, To))
						{
							mm.Subject = Subject;
							string body = string.Empty;
							Directory.GetCurrentDirectory();
							using (StreamReader reader = new StreamReader("C:\\LockedMail\\TibsNotifyMailTemplate.html"))
							{
								body = reader.ReadToEnd();
							}
							body = body.Replace("{Mail_Body}", MailBody);
							if (!string.IsNullOrEmpty(Cc))
							{
								mm.CC.Add(Cc);
							}
							mm.Body = body;
							mm.IsBodyHtml = true;
							SmtpClient smtp = new SmtpClient()
							{
								Host = SmtpServer,
								EnableSsl = true
							};
							NetworkCredential credentials = new NetworkCredential()
							{
								UserName = EmailAddress,
								Password = Password
							};
							smtp.UseDefaultCredentials = true;
							smtp.Credentials = credentials;
							smtp.Port = Port;
							mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
							smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
							try
							{
								this.WriteToFile("Error 2");
								smtp.Send(mm);
							}
							catch (SmtpFailedRecipientsException smtpFailedRecipientsException)
							{
								SmtpFailedRecipientsException ex = smtpFailedRecipientsException;
								this.WriteToFile("Catch 1");
								for (int i = 0; i < (int)ex.InnerExceptions.Length; i++)
								{
									SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
									if ((status == SmtpStatusCode.MailboxBusy ? false : status != SmtpStatusCode.MailboxUnavailable))
									{
										this.WriteToFile(string.Concat("Failed to deliver message to {0}", ex.InnerExceptions[i].FailedRecipient));
										SqlConnection conFinal4 = new SqlConnection(constrj1);
										conFinal4.Open();
										string qurfinal4 = string.Concat(new object[] { "update dbo.EmailDetail set IsSent='", true.ToString(), "',EmailStatus='", ex.InnerExceptions[i].FailedRecipient, "',LastModificationTime='", DateTime.Now, "' where Id='", Convert.ToInt32(datarow1["Id"]), "' " });
										(new SqlCommand(qurfinal4, conFinal4)).ExecuteNonQuery();
										conFinal4.Close();
									}
									else
									{
										this.WriteToFile(string.Concat("Failed to deliver message to {0}", ex.InnerExceptions[i].FailedRecipient));
										SqlConnection conFinal4 = new SqlConnection(constrj1);
										conFinal4.Open();
										string qurfinal4 = string.Concat(new object[] { "update dbo.EmailDetail set IsSent='", true.ToString(), "',EmailStatus='", ex.InnerExceptions[i].FailedRecipient, "',LastModificationTime='", DateTime.Now, "' where Id='", Convert.ToInt32(datarow1["Id"]), "' " });
										(new SqlCommand(qurfinal4, conFinal4)).ExecuteNonQuery();
										conFinal4.Close();
									}
								}
							}
							catch (Exception exception)
							{
								Exception ex = exception;
								this.WriteToFile(string.Concat("Exception caught in RetryIfBusy(): {0}", ex.ToString()));
								SqlConnection conFinal3 = new SqlConnection(constrj1);
								conFinal3.Open();
								string qurfinal3 = string.Concat(new object[] { "update dbo.EmailDetail set IsSent='", true.ToString(), "',EmailStatus='", ex.Message, "',LastModificationTime='", DateTime.Now, "' where Id='", Convert.ToInt32(datarow1["Id"]), "' " });
								(new SqlCommand(qurfinal3, conFinal3)).ExecuteNonQuery();
								conFinal3.Close();
							}
							this.WriteToFile(string.Concat("Email sent successfully to: ", From, " ", To));
							SqlConnection conFinal = new SqlConnection(constrj1);
							conFinal.Open();
							string qurfinal = string.Concat(new object[] { "update dbo.EmailDetail set IsSent='", true.ToString(), "',EmailStatus='Sent',LastModificationTime='", DateTime.Now, "' where Id='", Convert.ToInt32(datarow1["Id"]), "' " });
							(new SqlCommand(qurfinal, conFinal)).ExecuteNonQuery();
							conFinal.Close();
						}
					}
					catch (Exception exception1)
					{
						Exception ex = exception1;
						this.WriteToFile("Catch 2");
						this.WriteToFile(string.Concat("Error: ", ex.Message));
						SqlConnection conFinal2 = new SqlConnection(constrj1);
						conFinal2.Open();
						string qurfinal2 = string.Concat(new object[] { "update dbo.EmailDetail set IsSent='", true.ToString(), "',EmailStatus='Failed',LastModificationTime='", DateTime.Now, "' where Id='", Convert.ToInt32(datarow1["Id"]), "' " });
						(new SqlCommand(qurfinal2, conFinal2)).ExecuteNonQuery();
						conFinal2.Close();
					}
				}
				mainConnectionj1.Close();
			}
			catch (Exception exception2)
			{
				Exception ex = exception2;
				this.WriteToFile(string.Concat("Tibs Notify Service Job1 Error on: ", ex.Message, ex.StackTrace));
			}
		}

		private void WriteToFile(string text)
		{
			using (StreamWriter writer = new StreamWriter("C:\\LockedMail\\TibsNotifyJob2Log.txt", true))
			{
				DateTime now = DateTime.Now;
				writer.WriteLine(string.Format(text, now.ToString("dd/MM/yyyy hh:mm:ss tt")));
				writer.Close();
			}
		}
	}
}
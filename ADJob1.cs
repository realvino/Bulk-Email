using Quartz;
using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace WindowsServiceCS
{
	public class ADJob1 : IJob
	{
		public ADJob1()
		{
		}

		public void Execute(IJobExecutionContext context)
		{
			try
			{
				string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
				SqlConnection mainConnection = new SqlConnection(constr);
				mainConnection.Open();
				string Server = string.Empty;
				string Database = string.Empty;
				string Username = string.Empty;
				string Password = string.Empty;
				DataTable db1 = new DataTable();
				string query1 = "select * from [dbo].[NotificationManager] where Id = 1";
				using (SqlConnection con1 = new SqlConnection(constr))
				{
					using (SqlCommand cmd1 = new SqlCommand(query1))
					{
						cmd1.Connection = con1;
						using (SqlDataAdapter sda = new SqlDataAdapter(cmd1))
						{
							sda.Fill(db1);
						}
					}
				}
				foreach (DataRow datarow1 in db1.Rows)
				{
					Server = datarow1["Server"].ToString();
					Database = datarow1["Database"].ToString();
					Username = datarow1["Username"].ToString();
					Password = datarow1["Password"].ToString();
				}
				string myConnection = string.Concat(new string[] { "Data Source=", Server, ";Initial Catalog=", Database, ";User ID=", Username, ";Password=", Password });
				SqlConnection tempConnection = new SqlConnection(myConnection);
				DataTable db2 = new DataTable();
				bool flag = true;
				string query2 = string.Concat("select * from [dbo].[EmailNotification] where MailStatus='", flag.ToString(), "' and Duration is not null and IsDeleted = 'false'");
				using (SqlConnection con2 = new SqlConnection(constr))
				{
					using (SqlCommand cmd2 = new SqlCommand(query2))
					{
						cmd2.Connection = con2;
						using (SqlDataAdapter sda = new SqlDataAdapter(cmd2))
						{
							sda.Fill(db2);
						}
					}
				}
				try
				{
					foreach (DataRow datarow2 in db2.Rows)
					{
						this.WriteToFile("Tibs Notify Service: Select query from database success");
						string Id = datarow2["Id"].ToString();
						string Title = datarow2["Title"].ToString();
						string Query = datarow2["Query"].ToString();
						string ToEmailAddress = datarow2["ToEmailAddress"].ToString();
						string CcEmailAddress = datarow2["CcEmailAddress"].ToString();
						string ScheduleTime = datarow2["ScheduleTime"].ToString();
						string MailSubject = datarow2["MailSubject"].ToString();
						string MailBody = datarow2["MailBody"].ToString();
						string Duration = datarow2["Duration"].ToString();
						string PriColumn = datarow2["PriColumn"].ToString();
						string DatColumn = datarow2["DatColumn"].ToString();
						string IdColumn = datarow2["IdColumn"].ToString();
						string DateTimeColumn = datarow2["DateTiColumn"].ToString();
						int rate = int.Parse(Duration);
						string IdCoUp = string.Empty;
						string DaCoUp = string.Empty;
						string newIdCol = string.Empty;
						string newDatCol = string.Empty;
						this.WriteToFile(string.Concat("Id:", Id));
						this.WriteToFile(string.Concat("Running:", Query));
						newIdCol = (IdColumn == string.Empty ? string.Empty : string.Concat("a.", PriColumn, ">", IdColumn));
						newDatCol = (DateTimeColumn == string.Empty ? string.Empty : string.Concat(new string[] { "a.", DatColumn, "> '", DateTimeColumn, "'" }));
						DateTime dateTime = Convert.ToDateTime(ScheduleTime);
						DateTime TimeRun = dateTime.AddMinutes((double)rate);
						DateTime TimeNow = DateTime.Now;
						this.WriteToFile(string.Concat("TimeRun:", TimeRun));
						this.WriteToFile(string.Concat("TimeNow:", TimeNow));
						if (TimeNow > TimeRun)
						{
							string query3 = string.Empty;
							DataTable db3 = new DataTable();
							query3 = Query;
							if ((newIdCol != string.Empty ? false : newDatCol == string.Empty))
							{
								query3 = Query;
							}
							else if ((newIdCol == string.Empty ? false : newDatCol != string.Empty))
							{
								query3 = string.Concat(new string[] { "select * from (", query3, ")a Where ", newIdCol, " and ", newDatCol });
							}
							else if (newIdCol != string.Empty)
							{
								query3 = string.Concat("select * from (", query3, ")a Where ", newIdCol);
							}
							else if (newDatCol != string.Empty)
							{
								query3 = string.Concat("select * from (", query3, ")a Where ", newDatCol);
							}
							this.WriteToFile(string.Concat("Select query :", query3));
							using (SqlConnection con3 = new SqlConnection(myConnection))
							{
								using (SqlCommand cmd3 = new SqlCommand(query3))
								{
									cmd3.Connection = con3;
									using (SqlDataAdapter sda = new SqlDataAdapter(cmd3))
									{
										sda.Fill(db3);
									}
								}
							}
							try
							{
								foreach (DataRow datarow3 in db3.Rows)
								{
									this.WriteToFile("Tibs Notify Service: Select customer query from there database success");
									string ToEm = string.Empty;
									string CcEm = string.Empty;
									string SuEm = string.Empty;
									string BoEm = string.Empty;
									string IdCo = string.Empty;
									string DaCo = string.Empty;
									IdCoUp = string.Empty;
									DaCoUp = string.Empty;
									if (PriColumn != null)
									{
										if (IdColumn == string.Empty)
										{
											IdCoUp = "null";
										}
										else
										{
											IdCo = PriColumn;
											if (datarow3.Table.Columns.Contains(IdCo))
											{
												string IdCoDet = datarow3[IdCo ?? "" ?? ""].ToString();
												IdCoUp = string.Concat("'", IdCoDet, "'");
											}
										}
									}
									if (DatColumn != null)
									{
										if (DateTimeColumn == string.Empty)
										{
											DaCoUp = null;
										}
										else
										{
											DaCo = DatColumn;
											if (datarow3.Table.Columns.Contains(DaCo))
											{
												DaCoUp = datarow3[DaCo ?? "" ?? ""].ToString();
											}
										}
									}
									if (ToEmailAddress != null)
									{
										ToEm = ToEmailAddress;
										string[] output1 = (
											from Match m in Regex.Matches(ToEm, "\\[(.+?)\\]")
											select m.Groups[1].Value).ToArray<string>();
										for (int i = 0; i < output1.Count<string>(); i++)
										{
											string detail1 = output1[i];
											string de1 = string.Concat("[", detail1, "]");
											if (!datarow3.Table.Columns.Contains(detail1))
											{
												ToEm = ToEm.Replace(de1, "");
											}
											else
											{
												string ToAdd = datarow3[detail1 ?? "" ?? ""].ToString();
												ToEm = ToEm.Replace(de1, ToAdd);
											}
										}
									}
									if (CcEmailAddress != null)
									{
										CcEm = CcEmailAddress;
										string[] output2 = (
											from Match m in Regex.Matches(CcEm, "\\[(.+?)\\]")
											select m.Groups[1].Value).ToArray<string>();
										for (int i = 0; i < output2.Count<string>(); i++)
										{
											string detail2 = output2[i];
											string de2 = string.Concat("[", detail2, "]");
											if (!datarow3.Table.Columns.Contains(detail2))
											{
												CcEm = CcEm.Replace(de2, "");
											}
											else
											{
												string CcAdd = datarow3[detail2 ?? "" ?? ""].ToString();
												CcEm = CcEm.Replace(de2, CcAdd);
											}
										}
									}
									if (MailSubject != null)
									{
										SuEm = MailSubject;
										string[] output3 = (
											from Match m in Regex.Matches(SuEm, "\\[(.+?)\\]")
											select m.Groups[1].Value).ToArray<string>();
										for (int i = 0; i < output3.Count<string>(); i++)
										{
											string detail3 = output3[i];
											string de3 = string.Concat("[", detail3, "]");
											if (!datarow3.Table.Columns.Contains(detail3))
											{
												SuEm = SuEm.Replace(de3, "");
											}
											else
											{
												string MasuAdd = datarow3[detail3 ?? "" ?? ""].ToString();
												SuEm = SuEm.Replace(de3, MasuAdd);
											}
										}
									}
									if (MailBody != null)
									{
										BoEm = MailBody;
										string[] output4 = (
											from Match m in Regex.Matches(BoEm, "\\[(.+?)\\]")
											select m.Groups[1].Value).ToArray<string>();
										for (int i = 0; i < output4.Count<string>(); i++)
										{
											string detail4 = output4[i];
											string de4 = string.Concat("[", detail4, "]");
											if (!datarow3.Table.Columns.Contains(detail4))
											{
												BoEm = BoEm.Replace(de4, "");
											}
											else
											{
												string MaboAdd = datarow3[detail4 ?? "" ?? ""].ToString();
												BoEm = BoEm.Replace(de4, MaboAdd);
											}
										}
									}
									string query4 = string.Concat(new object[] { " INSERT INTO [dbo].[EmailDetail] ([EmailNotificationId],[Title],[Subject],[MailBody],[ToEmailAddress],[CcEmailAddress],[IsSent],[IsDeleted],[CreationTime]) VALUES(", Id, ",'", Title, "','", SuEm, "','", BoEm, "','", ToEm, "','", CcEm, "','false','false','", DateTime.Now, "')" });
									(new SqlCommand(query4, mainConnection)).ExecuteNonQuery();
								}
							}
							catch (Exception exception)
							{
								Exception ex = exception;
								this.WriteToFile(string.Concat("Tibs Notify Service Job1 Error on: ", ex.Message, ex.StackTrace));
							}
							if (db3.Rows.Count != 0)
							{
								if (DaCoUp != string.Empty)
								{
									DaCoUp = string.Concat("'", DaCoUp, "'");
								}
								if (DaCoUp == "''")
								{
									DaCoUp = "null";
								}
							}
							else
							{
								IdCoUp = (IdColumn == string.Empty ? "null" : string.Concat("'", IdColumn, "'"));
								DaCoUp = (DateTimeColumn == string.Empty ? "null" : string.Concat("'", DateTimeColumn, "'"));
								if (DaCoUp == "''")
								{
									DaCoUp = "null";
								}
							}
							string query5 = string.Concat(new string[] { " Update  EmailNotification set ScheduleTime = '", TimeNow.ToString("MM/dd/yyyy HH:mm:ss"), "', IdColumn = ", IdCoUp, ", DateTiColumn = ", DaCoUp, " where Id = ", Id, ";" });
							(new SqlCommand(query5, mainConnection)).ExecuteNonQuery();
							this.WriteToFile(string.Concat("Next Schedule Time :", TimeNow));
							this.WriteToFile(string.Concat("Next Update Query :", query5));
						}
					}
				}
				catch (Exception exception1)
				{
					Exception ex = exception1;
					this.WriteToFile(string.Concat("Tibs Notify Service Job1 Error on: ", ex.Message, ex.StackTrace));
				}
				mainConnection.Close();
			}
			catch (Exception exception2)
			{
				Exception ex = exception2;
				this.WriteToFile(string.Concat("Tibs Notify Service Job1 Error on: ", ex.Message, ex.StackTrace));
			}
		}

		private void WriteToFile(string text)
		{
			using (StreamWriter writer = new StreamWriter("C:\\LockedMail\\TibsNotifyJob1Log.txt", true))
			{
				DateTime now = DateTime.Now;
				writer.WriteLine(string.Format(text, now.ToString("dd/MM/yyyy hh:mm:ss tt")));
				writer.Close();
			}
		}
	}
}
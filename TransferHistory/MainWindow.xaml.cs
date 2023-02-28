using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.IO;
using System.Reflection;
using ClearScada.Client;


namespace TransferHistory
{
	// Per historic update
	class HisRecord
	{
		public DateTimeOffset RecordTime;
		public Double Value;
		public long Quality;
		public long Reason;
	}
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private string exePath;
		public MainWindow()
		{
			InitializeComponent();
			exePath = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			LoadINIFile(exePath + "\\TransferHistory.ini");
		}
		// Options and selections
		private string[] SelectedItems;
		private void LoadINIFile(string filename)
		{
			// INI file is located where this .exe is located
			if (!File.Exists(filename))
			{
				// No defaults file
				return;
			}
			SelectedItems = System.IO.File.ReadAllLines(filename);
			if (SelectedItems.Length >= 6)
			{
				ServerAddress.Text = SelectedItems[0];
				ServerPort.Text = SelectedItems[1];
				UserName.Text = SelectedItems[2];
				FileFolder.Text = SelectedItems[3];
				NameFilter.Text = SelectedItems[4];
				StartDate.SelectedDate = DateTime.Parse(SelectedItems[5]);
				EndDate.SelectedDate = DateTime.Parse(SelectedItems[6]);
				DateFormatString.Text = SelectedItems[7];
				// When the points are read, activate selections based on remaining data, skip a few spare & start at 10
				// Set focus on password
				password.Focus();
			}
		}
		private void SaveINIFile(string filename)
		{
			using (var tw = System.IO.File.CreateText(filename))
			{
				tw.WriteLine(ServerAddress.Text);
				tw.WriteLine(ServerPort.Text);
				tw.WriteLine(UserName.Text);
				tw.WriteLine(FileFolder.Text);
				tw.WriteLine(NameFilter.Text);
				tw.WriteLine(StartDate.SelectedDate);
				tw.WriteLine(EndDate.SelectedDate);
				tw.WriteLine(DateFormatString.Text);
				tw.WriteLine("");
				tw.WriteLine("");
				foreach (var s in listBox.SelectedItems)
				{
					tw.WriteLine(s);
				}
			}
		}

		// When enter key pressed in password field the points are read:
		private void OnKeyDownHandler(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Return)
			{
				Read_Points_Click(null, null);
			}
		}

		ClearScada.Client.Simple.Connection SimpleConnection;
		ClearScada.Client.Advanced.IServer AdvConnection;
		private void Read_Points_Click(object sender, RoutedEventArgs e)
		{
			// Connect to database
			ServerNode node = new ClearScada.Client.ServerNode(ServerAddress.Text, int.Parse(ServerPort.Text));
			SimpleConnection = new ClearScada.Client.Simple.Connection("Utility");
			try
			{
				SimpleConnection.Connect(node);
				AdvConnection = SimpleConnection.Server;
			}
			catch (Exception err)
			{
				MessageBox.Show("Cannot connect. " + err.Message);
				return;
			}
			try
			{
				SimpleConnection.LogOn(UserName.Text, password.SecurePassword);
			}
			catch (Exception err)
			{
				MessageBox.Show("Cannot log in. " + err.Message);
				return;
			}
			// Query points

			string sql = "SELECT O.FullName As OName, O.TypeDesc FROM CHistory H INNER JOIN CDBObject O ON H.Id=O.Id WHERE OName LIKE '" + NameFilter.Text + "' ORDER BY OName";
			ClearScada.Client.Advanced.IQuery serverQuery = AdvConnection.PrepareQuery(sql, new ClearScada.Client.Advanced.QueryParseParameters());
			ClearScada.Client.Advanced.QueryResult queryResult = serverQuery.ExecuteSync(new ClearScada.Client.Advanced.QueryExecuteParameters());
			if (queryResult.Status == ClearScada.Client.Advanced.QueryStatus.Succeeded || queryResult.Status == ClearScada.Client.Advanced.QueryStatus.NoDataFound)
			{
				if (queryResult.Rows.Count > 0)
				{
					Dispatcher.Invoke(new Action(() =>
					{
						listBox.Items.Clear();
					}));
					IEnumerator<ClearScada.Client.Advanced.QueryRow> rows = queryResult.Rows.GetEnumerator();
					while (rows.MoveNext())
					{
						Dispatcher.Invoke(new Action(() =>
						{
							listBox.Items.Add((string)rows.Current.Data[0] + " (" + (string)rows.Current.Data[1] + ")");
						}));
					}
				}
				else
				{
					MessageBox.Show("No rows found matching the search");
					return;
				}
			}
			else
			{
				MessageBox.Show("Database query error");
				return;
			}
			serverQuery.Dispose();

			// Loaded selections
			if (SelectedItems != null && SelectedItems.Length >= 10)
			{
				for (int i = 10; i < SelectedItems.Length; i++)
				{
					if (listBox.Items.Contains(SelectedItems[i]))
					{
						listBox.SelectedItems.Add(SelectedItems[i]);
					}
				}
			}
		}

		private void SelectNone_Click(object sender, RoutedEventArgs e)
		{
			listBox.SelectedIndex = -1;
		}

		private void SelectAll_Click(object sender, RoutedEventArgs e)
		{
			listBox.SelectAll();
		}

		private void Quit_Click(object sender, RoutedEventArgs e)
		{
			this.Close();
		}

		async private void Export_Click(object sender, RoutedEventArgs e)
		{
			SaveINIFile(exePath + "\\TransferHistory.ini");

			int pointnumber = 0;
			int totalrecords = 0;

			string startDateFormatted = ((DateTime)(StartDate.SelectedDate)).ToString("yyyy-MM-dd HH:mm:ss.fff");
			string endDateFormatted = ((DateTime)(EndDate.SelectedDate)).ToString("yyyy-MM-dd HH:mm:ss.fff");
			string SQLConstraint = "\"RecordTime\" BETWEEN {TS '" + startDateFormatted + "'} AND {TS '" + endDateFormatted + "'}";

			foreach (string pointnamedesc in listBox.SelectedItems)
			{
				string pointname = pointnamedesc.Substring(0, pointnamedesc.LastIndexOf('(') - 1);
				string fn = FileFolder.Text + "\\" + pointname.Replace('*', '$') + ".txt";

				string columnNames = "\"RecordTime\", \"ValueAsReal\", \"StateDesc\", \"QualityDesc\", \"ReasonDesc\"";
				string SQL = "SELECT " + columnNames + " FROM CDBHISTORIC H INNER JOIN CDBObject O ON H.Id=O.Id WHERE O.FullName = '" + pointname + "' And " + SQLConstraint + " Order By \"RecordTime\" ASC";

				ClearScada.Client.Advanced.IQuery serverQuery = AdvConnection.PrepareQuery(SQL, new ClearScada.Client.Advanced.QueryParseParameters());
				ClearScada.Client.Advanced.QueryResult queryResult = serverQuery.ExecuteSync(new ClearScada.Client.Advanced.QueryExecuteParameters());

				if (queryResult.Status == ClearScada.Client.Advanced.QueryStatus.Succeeded || queryResult.Status == ClearScada.Client.Advanced.QueryStatus.NoDataFound)
				{
					if (queryResult.Rows.Count > 0)
					{
						// Open file
						using (var his = System.IO.File.CreateText(fn))
						{
							pointnumber++;
							// Write header
							his.WriteLine(columnNames.Replace(',', '\t'));

							IEnumerator<ClearScada.Client.Advanced.QueryRow> rows = queryResult.Rows.GetEnumerator();
							int rownumber = 0;
							while (rows.MoveNext())
							{
								if (rownumber % 1000 == 0)
								{
									await Dispatcher.BeginInvoke(new Action(() =>
									{
										Progress.Text = "Row " + rownumber.ToString() + " of " + queryResult.Rows.Count.ToString() + " for " + pointname;
									}));
									await Task.Delay(1);
								}
								rownumber++;

								string nextline = "";
								foreach (var entry in rows.Current.Data)
								{
									//Console.WriteLine(entry.GetType().Name);

									switch (entry.GetType().Name)
									{
										case "String":
											nextline += "\"" + (string)entry + "\"\t";
											break;
										case "DateTime":
											nextline += ((DateTime)entry).ToString(DateFormatString.Text) + "\t";
											break;
										case "DateTimeOffset":
											nextline += ((DateTimeOffset)entry).ToString(DateFormatString.Text) + "\t";
											break;
										default:
											nextline += entry.ToString() + "\t";
											break;
									}
								}
								// Remove last tab
								his.WriteLine(nextline.Substring(0, nextline.Length - 1));
							}
							totalrecords += rownumber;
						}
					}
					else
					{
						// No rows found, do not output file for this point
					}
				}
				else
				{
					await Dispatcher.BeginInvoke(new Action(() =>
					{
						Progress.Text = "Database query error";
					}));
					return;
				}
				serverQuery.Dispose();
			}
			await Dispatcher.BeginInvoke(new Action(() =>
			{
				Progress.Text = $"Written {pointnumber} files and {totalrecords} records.";
			}));
		}

		async private void Import_Click(object sender, RoutedEventArgs e)
		{
			SaveINIFile(exePath + "\\TransferHistory.ini");

			int pointcount = 0;
			int totalrecords = 0;
			int totalrecordsthispoint = 0;

			// Connect to database
			ServerNode node = new ClearScada.Client.ServerNode(ServerAddress.Text, int.Parse(ServerPort.Text));
			SimpleConnection = new ClearScada.Client.Simple.Connection("Utility");
			SimpleConnection.Connect(node);
			AdvConnection = SimpleConnection.Server;
			try
			{
				SimpleConnection.LogOn(UserName.Text, password.SecurePassword);
			}
			catch (Exception err)
			{
				await Dispatcher.BeginInvoke(new Action(() =>
				{
					Progress.Text = "Cannot log in. " + err.Message;
				}));
				return;
			}

			// Read files in the input directory
			var files = Directory.GetFiles(FileFolder.Text);
			foreach (var filename in files)
			{
				var ObjName = Path.GetFileNameWithoutExtension(filename).Replace('$', '*'); // Replace(Left(Fin, InStrRev(Fin, ".") - 1), "$", "*");
				var PointObj = SimpleConnection.GetObject(ObjName);
				if (PointObj != null) // May also check this is a point/accumulator with historic storage
				{
					pointcount++;
					var lines = File.ReadAllLines(filename);
					var HisRecords = new List<HisRecord>();
					// Format: columnNames = "\"RecordTime\", \"ValueAsReal\", \"StateDesc\", \"QualityDesc\", \"ReasonDesc\"";
					foreach (var line in lines)
					{
						// If line contains RecordTime then ignore it - it's a header
						if (line.Contains("RecordTime")) continue;
						// Split line into fields
						var fields = line.Split('\t');
						// Handle CSV as an alternative
						if (line.Length == 1)
						{
							fields = line.Split(',');

						}
						// Remove " from fields
						for (int i = 0; i < fields.Length; i++)
						{
							if (fields[i].StartsWith("\"") && fields[i].EndsWith("\""))
							{
								fields[i] = fields[i].Substring(1, fields[i].Length - 2);
							}
						}
						// check value is a date
						var Record = new HisRecord();
						if (DateTimeOffset.TryParse(fields[0], out Record.RecordTime))
						{
							// Get sample value
							if (Double.TryParse(fields[1], out Record.Value))
							{
								totalrecordsthispoint++;

								Record.Quality = DecodeQuality(fields[3]);
								Record.Reason = DecodeReason(fields[4]);

								// If the reason is 8 then we use a modification procedure, else use array import import
								if (Record.Reason != 8)
								{
									HisRecords.Add(Record);
									if (HisRecords.Count >= 5000) // Batch Size
									{
										WriteHistory(HisRecords, PointObj);
										await Dispatcher.BeginInvoke(new Action(() =>
										{
											Progress.Text = $"Wrote {totalrecordsthispoint} records to {ObjName}";
										}));
										await Task.Delay(1);
									}
								}
								else
								{
									var args = new object[3];
									args[0] = Record.RecordTime;
									args[1] = Record.Value;
									args[2] = Record.Quality;
									args[3] = ""; // Event log comment
									PointObj.InvokeMethod("Historic.ModifyValue", args);
								}
							}
						} // else can't read the date
					}
					// Remaining records
					if (HisRecords.Count > 0)
					{
						WriteHistory(HisRecords, PointObj);
						await Dispatcher.BeginInvoke(new Action(() =>
						{
							Progress.Text = $"Wrote {totalrecordsthispoint} records to {ObjName}";
						}));
						await Task.Delay(1);
					}
					totalrecords += totalrecordsthispoint;
				} // else point does not exist
			}
			await Dispatcher.BeginInvoke(new Action(() =>
			{
				Progress.Text = $"Complete. Wrote {totalrecords} records to {pointcount} points.";
			}));
		}
		void WriteHistory( List<HisRecord> HisRecords, ClearScada.Client.Simple.DBObject PointObj)
		{
			// Convert our records into four separate arrays
			var t = new DateTimeOffset[HisRecords.Count];
			var v = new Double[HisRecords.Count];
			var q = new long[HisRecords.Count];
			var r = new long[HisRecords.Count];
			long i = 0;
			foreach (var rec in HisRecords)
			{
				t[i] = rec.RecordTime;
				v[i] = rec.Value;
				q[i] = rec.Quality;
				r[i] = rec.Reason;
				i++;
			}
			var args = new object[4];
			args[0] = v; // Value first
			args[1] = t;
			args[2] = q;
			args[3] = r;
			PointObj.InvokeMethod("Historic.LoadDataValues", args);
			// Empty the list
			HisRecords.Clear();
		}
		int DecodeQuality(string Name)
		{
			switch (Name)
			{
				case "SubNormal":
					return 88;
				case "Good":
					return 192;
				case "Bad":
					return 0;
				case "Local Override":
					return 216;
				case "Engineering Units Exceeded":
					return 84;
				default: // "Good":
					return 192;
			}
		}
		int DecodeReason(string Name)
		{
			switch (Name)
			{
				case "Current Data":
					return 0;
				case "CDR":
					return 0;
				case "Value Change":
					return 1;
				case "State Change":
					return 2;
				case "Report":
					return 3;
				case "Timed Report":
					return 3;
				case "End of Period":
					return 4;
				case "End of Period Reset":
					return 5;
				case "Override":
					return 6;
				case "Release":
					return 7;
				case "Release Override":
					return 7;
				case "Modified/Inserted":
					return 8;
				case "Modified":
					return 8;
				case "Inserted":
					return 8;
				default: // "Current Data":
					return 0;
			}
		}

		private void listBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			// Count and state # selected
			Progress.Text = $"{listBox.SelectedItems.Count} points selected.";
		}
	}
}
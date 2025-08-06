using ClearScada.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
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
using System.Windows.Shapes;
using Path = System.IO.Path;


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
            //File.Delete(exePath + "\\TransferHistory.ini"); 
            StartDate.SelectedDate = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            EndDate.SelectedDate = DateTime.UtcNow;

            LoadINIFile(exePath + "\\TransferHistory.ini");
            DateFormatCombo.Items.Add("Select");
            DateFormatCombo.SelectedIndex = 0;
            DateFormatCombo.Items.Add("Simple");
            DateFormatCombo.Items.Add("ANSI (DataFile Export)");
            DateFormatCombo.Items.Add("ISO8601");
        }

        // Date preformats
        private void DateFormatCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DateFormatCombo.SelectedIndex > 0)
            {
                switch (DateFormatCombo.SelectedIndex)
                {
                    case 1:
                        DateFormatString.Text = "yyyy-MM-dd HH:mm:ss.fff";
                        break;
                    case 2:
                        DateFormatString.Text = "yyyy,MM,dd,HH,mm,ss";
                        break;
                    case 3:
                        DateFormatString.Text = "yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fffffffzzz";
                        break;
                    default:
                        break;
                }
            }
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
				if (SelectedItems[8].ToLower() == "true")
				{
					UTCBox.IsChecked = true;
				}
				// When the points are read, activate selections based on remaining data, skip a few spare & start at 10
				// Set focus on password
				password.Focus();
			}
		}

		private void SaveINIFile(string filename)
		{
			using (var tw = System.IO.File.CreateText(filename))
			{
				tw.WriteLine(ServerAddress.Text); //0
				tw.WriteLine(ServerPort.Text); //1
				tw.WriteLine(UserName.Text); //2
				tw.WriteLine(FileFolder.Text); //3
				tw.WriteLine(NameFilter.Text); //4
				tw.WriteLine(StartDate.SelectedDate); //5
				tw.WriteLine(EndDate.SelectedDate); //6
				tw.WriteLine(DateFormatString.Text); //7
				tw.WriteLine(((bool)(UTCBox.IsChecked)).ToString()); //8
				tw.WriteLine(""); //9
				foreach (var s in listBox.SelectedItems) //10...
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
			SaveINIFile(exePath + "\\TransferHistory.ini");

            // Connect to database
            try
            {
                ServerNode node = new ClearScada.Client.ServerNode(ServerAddress.Text, int.Parse(ServerPort.Text));
				SimpleConnection = new ClearScada.Client.Simple.Connection("Utility");
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
            // Mackenzie Wilkins: Added support for Exporting as ANSI file
            bool Export_ANSI = false;
            Encoding encoding = Encoding.UTF8;
            if (sender.Equals(Export_Copy))
            {
                //MessageBox.Show("ANSI");
                DateFormatCombo.SelectedIndex = 2;
                DateFormatCombo_SelectionChanged(null, null);
                encoding = Encoding.GetEncoding(1252);
                Export_ANSI = true;
				UTCBox.IsChecked = true;
            }

            System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show(
                "When importing, the \"Export (As DataFile)\" feature is much faster than the normal \"Export\" option. The Datafile option does require the file to exist on the GeoSCADA Server Host.\n\nClick Yes if you want to proceed with the \'" + (Export_ANSI?"DataFile":"slower")+"\' option.", // The message to display
				"Will you be on the GeoSCADA Server Host?",           // The title of the message box
                System.Windows.Forms.MessageBoxButtons.YesNo,  // Specifies Yes and No buttons
				System.Windows.Forms.MessageBoxIcon.Question   // Optional: Adds a question icon
			);
            if (result == System.Windows.Forms.DialogResult.No)
            {
                // Code to execute if the user clicks "No"
                await Dispatcher.BeginInvoke(new Action(() =>
                {
                    Progress.Text = "Export Cancelled";
                }));
                return;
            }
            await Dispatcher.BeginInvoke(new Action(() =>
            {
                Progress.Text = "Prepare to export";
            }));
            await Task.Delay(1);

            SaveINIFile(exePath + "\\TransferHistory.ini");

            int pointnumber = 0;
            int totalrecords = 0;
            string SQLConstraint = "";

            try
            {
                string startDateFormatted = ((DateTime)(StartDate.SelectedDate)).ToString("yyyy-MM-dd HH:mm:ss.fff");
                string endDateFormatted = ((DateTime)(EndDate.SelectedDate)).ToString("yyyy-MM-dd HH:mm:ss.fff");
                SQLConstraint = "\"RecordTime\" BETWEEN {TS '" + startDateFormatted + "'} AND {TS '" + endDateFormatted + "'}";
            }
            catch (Exception err)
            {
                System.Windows.MessageBox.Show("Issue with selected date " + err.Message);
                return;
            }


            if (!Directory.Exists(FileFolder.Text))
            {
                System.Windows.MessageBox.Show("Folder does not exist.");
                return;
            }

            foreach (string pointnamedesc in listBox.SelectedItems)
            {
                string pointname = pointnamedesc.Substring(0, pointnamedesc.LastIndexOf('(') - 1);
                string fn = FileFolder.Text + "\\" + pointname.Replace('*', '$') + ".txt";

                // if Exporting as ANSI, don't query additional data
                string columnNames = "\"RecordTime\", \"ValueAsReal\"" + (Export_ANSI ? "" : ", \"StateDesc\", \"QualityDesc\", \"ReasonDesc\"");
                string SQL = "SELECT " + columnNames + " FROM CDBHISTORIC H INNER JOIN CDBObject O ON H.Id=O.Id WHERE O.FullName = '" + pointname + "' And " + SQLConstraint + " Order By \"RecordTime\" ASC";

                ClearScada.Client.Advanced.IQuery serverQuery = AdvConnection.PrepareQuery(SQL, new ClearScada.Client.Advanced.QueryParseParameters());
                ClearScada.Client.Advanced.QueryResult queryResult = serverQuery.ExecuteSync(new ClearScada.Client.Advanced.QueryExecuteParameters());

                if (queryResult.Status == ClearScada.Client.Advanced.QueryStatus.Succeeded || queryResult.Status == ClearScada.Client.Advanced.QueryStatus.NoDataFound)
                {
                    if (queryResult.Rows.Count > 0)
                    {
                        // Open file
                        using (var his = new StreamWriter(fn, false, encoding))//System.IO.File.CreateText(fn))
                        {
                            pointnumber++;
                            // Write header if not Ansi
                            if (!Export_ANSI)
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

                                string delimiter = "\t";

                                if (Export_ANSI)
                                    delimiter = ",";

                                foreach (var entry in rows.Current.Data)
                                {
                                    //Console.WriteLine(entry.GetType().Name);

                                    switch (entry.GetType().Name)
                                    {
                                        case "String":
                                            nextline += "\"" + (string)entry + "\"" + delimiter;
                                            break;
                                        case "DateTime":
                                            nextline += ((DateTime)entry).ToString(DateFormatString.Text) + delimiter;
                                            break;
                                        case "DateTimeOffset":
                                            if ((bool)!UTCBox.IsChecked)
                                            {
                                                nextline += ((DateTimeOffset)entry).LocalDateTime.ToString(DateFormatString.Text) + delimiter;
                                            }
                                            else
                                            {
                                                nextline += ((DateTimeOffset)entry).UtcDateTime.ToString(DateFormatString.Text) + delimiter;
                                            }
                                            break;
                                        default:
                                            nextline += entry.ToString() + delimiter;
                                            break;
                                    }
                                }
                                // Remove last delimiter
                                his.WriteLine(nextline.Substring(0, nextline.Length - delimiter.Length));
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
            System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show(
                "The \"Import\" feature is much slower than the normal \"Import (As DataFile)\" option. If you have more than ~100 historic points, you have been advised NOT to use this method.\n\nClick Yes if you want to proceed.", // The message to display
                "Are you sure you want to do this?",           // The title of the message box
                System.Windows.Forms.MessageBoxButtons.YesNo,  // Specifies Yes and No buttons
                System.Windows.Forms.MessageBoxIcon.Question   // Optional: Adds a question icon
            );
            if (result == System.Windows.Forms.DialogResult.No)
            {
                // Code to execute if the user clicks "No"
                MessageBox.Show("Import Cancelled");
                return;
            }

            await Dispatcher.BeginInvoke(new Action(() =>
			{
				Progress.Text = "Prepare to import";
			}));
			await Task.Delay(1);

			SaveINIFile(exePath + "\\TransferHistory.ini");

			int pointcount = 0;
			int totalrecords = 0;
			int totalrecordsthispoint = 0;

            // Connect to database
            try
            {
                ServerNode node = new ClearScada.Client.ServerNode(ServerAddress.Text, int.Parse(ServerPort.Text));
				SimpleConnection = new ClearScada.Client.Simple.Connection("Utility");
				SimpleConnection.Connect(node);
				AdvConnection = SimpleConnection.Server;
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
					int rownum = 0;
					foreach (var line in lines)
					{
						rownum++;
						// If line contains RecordTime and row <= 2 then ignore it - it's a header
						if (line.Contains("RecordTime") && rownum <= 2) continue;
						// Split line into fields
						var fields = line.Split('\t');
						// Handle CSV as an alternative
						if (fields.Length == 1)
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
							// We read a value in using Local time, but if UTC we need to add/change the offset
							// This is questionable but achieves the desired effect
							if ((bool)UTCBox.IsChecked)
							{
								Record.RecordTime = Record.RecordTime.Add(Record.RecordTime.Offset);
							}
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
										try
										{
											WriteHistory(HisRecords, PointObj);
											await Dispatcher.BeginInvoke(new Action(() =>
											{
												Progress.Text = $"Wrote {totalrecordsthispoint} records to {ObjName}";
											}));
											await Task.Delay(1);
										}
										catch (Exception ex)
										{
											await Dispatcher.BeginInvoke(new Action(() =>
											{
												Progress.Text = $"Error {ex.Message} writing to {ObjName}";
											}));
											await Task.Delay(1);
										}
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
							} // else can't read the value
							else
							{
								await Dispatcher.BeginInvoke(new Action(() =>
								{
									Progress.Text = $"Cannot read value: {fields[1]}";
								}));
								await Task.Delay(1);
							}
						} // else can't read the date
						else
						{
							await Dispatcher.BeginInvoke(new Action(() =>
							{
								Progress.Text = $"Cannot read date: {fields[0]}";
							}));
							await Task.Delay(1);
						}
					}
					// Remaining records
					if (HisRecords.Count > 0)
					{
						try
						{
							WriteHistory(HisRecords, PointObj);
							await Dispatcher.BeginInvoke(new Action(() =>
							{
								Progress.Text = $"Wrote {totalrecordsthispoint} records to {ObjName}";
							}));
							await Task.Delay(1);
						}
						catch (Exception ex)
						{
							await Dispatcher.BeginInvoke(new Action(() =>
							{
								Progress.Text = $"Error {ex.Message} writing to {ObjName}";
							}));
							await Task.Delay(1);
						}
					}
					totalrecords += totalrecordsthispoint;
				} // else point does not exist
			}
			await Dispatcher.BeginInvoke(new Action(() =>
			{
				Progress.Text = $"Complete. Wrote {totalrecords} records to {pointcount} points.";
			}));
		}

        async private void Import_Click_DataFile(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show(
				"The \"Import (As DataFile)\" feature is much faster than the normal \"Import\" option but requires the data was exported in the 'DataFile' format and to exist on the GeoSCADA Server Host.\n\nClick Yes if you want to proceed.", // The message to display
				"Are you on the GeoSCADA Server Host?",           // The title of the message box
				System.Windows.Forms.MessageBoxButtons.YesNo,  // Specifies Yes and No buttons
				System.Windows.Forms.MessageBoxIcon.Question   // Optional: Adds a question icon
			);
            if (result == System.Windows.Forms.DialogResult.No)
            {
                // Code to execute if the user clicks "No"
                MessageBox.Show("Import Cancelled");
                return;
            }

            await Dispatcher.BeginInvoke(new Action(() =>
            {
                Progress.Text = "Prepare to import";
            }));
            await Task.Delay(1);

            SaveINIFile(exePath + "\\TransferHistory.ini");

            int filecount = 0;
            int filetotal = 0;
			int skipped = 0;

            // Connect to database
            try
            {
                ServerNode node = new ClearScada.Client.ServerNode(ServerAddress.Text, int.Parse(ServerPort.Text));
                SimpleConnection = new ClearScada.Client.Simple.Connection("Utility");
                SimpleConnection.Connect(node);
                AdvConnection = SimpleConnection.Server;
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
				++filetotal;
                //MessageBox.Show(filename);
                var ObjName = Path.GetFileNameWithoutExtension(filename).Replace('$', '*'); // Replace(Left(Fin, InStrRev(Fin, ".") - 1), "$", "*");
                var PointObj = SimpleConnection.GetObject(ObjName);

				// May also check this is a point/accumulator with historic storage
				if (PointObj == null)
				{
					++skipped;
					continue;
				}
				// if file has header row (i.e. Contains "RecordTime"), it is not in DataFile Format and should be skipped
				var line = File.ReadLines(filename).First();
				if (line.Contains("RecordTime"))
				{
                    ++skipped;
                    continue;
				}

                try
                {
                    PointObj.InvokeMethod("Historic.LoadDataFile", files);
					++filecount;
                }
                catch (Exception err)
                {
                    MessageBox.Show("File load didn't work. Error: " + err.Message);
                    return;
                }
            }
            await Dispatcher.BeginInvoke(new Action(() =>
            {
                Progress.Text = $"Complete. Wrote {filecount} files to {filetotal} points." +(skipped==0 ? "" : "{skipped} points skipped due to file not being correct format or history not enabled on point.");
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
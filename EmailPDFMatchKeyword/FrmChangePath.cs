using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EmailPDFMatchKeyword
{
	public partial class FrmChangePath : Form
	{
		public FrmChangePath()
		{
			InitializeComponent();
			txtwritepath.Text = GetDriveFolderPath();
			txtsheetid.Text = GetSpreadsheetId();
		}
		private void btnpathsave_Click(object sender, EventArgs e)
		{
			try
			{
				
				string textwriteValue = txtwritepath.Text.Trim();

				if (string.IsNullOrEmpty(textwriteValue))
				{
					MessageBox.Show(this, "Please enter a drive path value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				string filePath1 = Path.Combine(Application.StartupPath, "basepath.txt");
				string sheetId = ExtractSpreadsheetId(txtsheetid.Text.Trim());

				if (string.IsNullOrEmpty(sheetId))
				{
					MessageBox.Show("Please enter SpreadsheetId");
					return;
				}
				UpdateSpreadsheetId(sheetId);
				File.WriteAllText(filePath1, textwriteValue);

				MessageBox.Show(this, "Path saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message);
			}
		}

		private void btncancel_Click(object sender, EventArgs e)
		{
			this.Close();
		}
		private static string GetDriveFolderPath()
		{
			string filePath = Path.Combine(
				Application.StartupPath,
				"basepath.txt"
			);

			// If txt file not exists
			if (!File.Exists(filePath))
			{
				return "";
			}

			// Read txt value
			string txtValue = File.ReadAllText(filePath).Trim();

			// If txt is empty → use appsettings
			if (string.IsNullOrEmpty(txtValue))
			{
				return "";
			}

			// Else use txt value
			return txtValue;
		}

		private void btn2Browse_Click(object sender, EventArgs e)
		{
			using (FolderBrowserDialog dialog = new FolderBrowserDialog())
			{
				dialog.Description = "Select folder to save documents";
				dialog.ShowNewFolderButton = true;

				// Set default path (optional)
				string currentPath = txtwritepath.Text;

				if (Directory.Exists(currentPath))
				{
					dialog.SelectedPath = currentPath;
				}

				if (dialog.ShowDialog() == DialogResult.OK)
				{
					// Set selected path to textbox
					txtwritepath.Text = dialog.SelectedPath;
				}
			}
		}

		private string GetSpreadsheetId()
		{
			try
			{
				string path = Path.Combine(Application.StartupPath, "appsettings.json");
				if (!File.Exists(path)) return "";

				dynamic json = JsonConvert.DeserializeObject(File.ReadAllText(path));
				return json["GoogleDrive"]["SpreadsheetId"];
			}
			catch
			{
				return "";
			}
		}

		private void UpdateSpreadsheetId(string newId)
		{
			try
			{
				string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..\\..\\..\\appsettings.json");
				path = Path.GetFullPath(path);

				if (!File.Exists(path))
				{
					MessageBox.Show("appsettings.json not found.");
					return;
				}
				string json = File.ReadAllText(path);

				json = Regex.Replace(
				json,
				"\"SpreadsheetId\"\\s*:\\s*\".*?\"",
				$"\"SpreadsheetId\": \"{newId}\""
					);

				File.WriteAllText(path, json);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Failed to update SpreadsheetId: " + ex.Message);
			}			
		}

		private string ExtractSpreadsheetId(string input)
		{
			var match = Regex.Match(input, @"/d/([^/]+)");
			return match.Success ? match.Groups[1].Value : input;
		}


	}
}

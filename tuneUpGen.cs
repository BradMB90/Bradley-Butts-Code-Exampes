using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Net;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;

namespace TuneUpGen
{
    public partial class tuneUpGen : Form
    {
        //connection string to our Dynamics DB
        private string connectionString = "Provider=SQLOLEDB.1;Password=S3rvice!;Persist Security Info=True;User ID=sa;Initial Catalog=GTG_MSCRM;Data Source=dynamics";
        OleDbConnection connection;
        private string selectedClient, clientName, clientPOC, clientStatus, clientIssues, clientArcVersion, clientOpportunites, clientSummary, reportAuthor, reportDate, filename, emailRecipient;

        //Function that handes the report generator's load and close events.
        public tuneUpGen()
        {
            InitializeComponent();
            Load += new EventHandler(tuneUpGen_Load);
            Closed += new EventHandler(tuneUpGen_Close);
        }

        /*UpdatePOC()
         * The following function updates the drop down box the the application window for POC. Once a user select's the client they wish to generate a report off of, this function runs
         * to update the combo box with only the POC associated with that client. That is achieved using a SQL statement to select all clients in the ContactBase, with a ParentCustomerID
         * of the selected client; selectedClient
         */
        private void UpdatePOC()
        {
            try
            {
                pocComboBox.Items.Clear();
                pocComboBox.Text = "";
                connection = new OleDbConnection(connectionString);
                connection.Open();
                pocComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                pocComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;
                OleDbCommand getClientPOC = new OleDbCommand("select FullName from GTG_MSCRM.dbo.ContactBase where ParentCustomerID = '" + selectedClient + "' order by FullName", connection);
                OleDbDataReader reader = getClientPOC.ExecuteReader();
                while (reader.Read())
                {
                    pocComboBox.Items.Add(reader[0].ToString());
                }
                connection.Close();
            }
            catch (Exception getPOC_ex)
            {
                MessageBox.Show(getPOC_ex.ToString(), "Error Generating AccountID");
            }
        }

        /*UpdateAssets()
         * The following function automatically launches when a user selects the client they wish to run a report off of. Instead of filling out a ComboBox, we have the asset information being
         * placed in a DataSet. This allows us to view their current assests and the maintenance dates on said assets once a client is selected. The sql statement being used to pull this information
         * also contains a WHERE clause to only show assets that have a maintenance due date that hasn't come up yet (GETDATE < MAINTENANCEDUE), or if the maintenance due date is only 3 months past due.
         * This is due to SunGard. Since clients don't pay us directly; we have to wait for SunGard to send us our portion of the check. This could take any where between 2-3 months.
         */
        private void UpdateAssets()
        {
            try
            {
                connection = new OleDbConnection(connectionString);
                connection.Open();
                string getAssets = "select new_name as Asset, new_maintenancedue as Due_Date from GTG_MSCRM.dbo.new_assetExtensionBase where new_accountid = '" + selectedClient + "' and (new_maintenancedue > GETDATE() or DATEDIFF(month, new_maintenancedue, GETDATE()) < 5)";
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(getAssets, connection);
                DataSet dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                assetData.ReadOnly = true;
                assetData.DataSource = dataSet.Tables[0];
            }
            catch(Exception assetPopulate_ex)
            {
                MessageBox.Show(assetPopulate_ex.ToString(), "Error Populating Asset List");
            }
        }

        /*testConnection()
         * The following function is used to make sure the connection string being used to connect to Dynamics is valid. If it fails to connect to Dynamics, it'll inform the user.
         */
        private void testConnection()
        {
            try
            {
                connection = new OleDbConnection(connectionString);
                connection.Open();
                connection.Close();
            }
            catch (Exception load_ex)
            {
                MessageBox.Show(load_ex.ToString(), "Error Connecting to Dynamics");
            }
        }

        /*addAuthorAndDate()
         * Once the application is launched, the following function will be called. The fields reportAuthor and reportDate are filled in by pulling the users Windows' Display Name
         * and the current date from the user's machine.
         */
        private void addAuthorAndDate()
        {
            try
            {
                reportAuthor = System.DirectoryServices.AccountManagement.UserPrincipal.Current.DisplayName;
                authorTextBox.Text = reportAuthor;
                reportDate = DateTime.Today.ToString("d");
                dateTextBox.Text = reportDate;
            }
            catch (Exception authorAndDate_ex)
            {
                MessageBox.Show(authorAndDate_ex.ToString(), "Error Generating Author Name and Date");
            }
        }

        /*tuneUpGen_Load()
         * The following function is enabled once the application window loads. It'll go through a DB connection test, and populate the reportAuthor and reportDate. Once it goes through
         * those steps, it'll open a connection to the Dynamics SQL DB. Once connected, a SQL query is ran to populate the combobox labeled clientComboBox. The SQL query is going to populate
         * the combo box with all active clients currently in our SQL DB. From there, the user is able to select a client and continue their report.
         */
        private void tuneUpGen_Load(object sender, System.EventArgs e)
        {
            try
            {
                testConnection();
                addAuthorAndDate();
                connection.Open();
                clientComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                clientComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;
                OleDbCommand fillComboBox = new OleDbCommand("SELECT AccountID, Name FROM [GTG_MSCRM].[dbo].[AccountBase] where StateCode = 0 order by Name", connection);
                OleDbDataReader reader = fillComboBox.ExecuteReader();
                while (reader.Read())
                {
                    
                    clientComboBox.Items.Add(reader[1].ToString());
                }
                connection.Close();
            }
            catch (Exception clientCombo_ex)
            {
                MessageBox.Show(clientCombo_ex.ToString(), "Error generating the client list");
            }           
        }

        /*tuneUpGen_Close()
         * Once the users selects to close the application, the connection to the Dynamics SQL DB is closed as well.
         */
        private void tuneUpGen_Close(object sender, System.EventArgs e)
        {
            OleDbConnection connection = new OleDbConnection(connectionString);

            try
            {
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error Closing Connection to Dynamics");
            }
        }

        /*clientComboBox_SelectedIndexChanged()
        * The following function changes the POC comboxbox and the Assets dataset based on the client selected in
        * the client combo box. 
        */
        private void clientComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                connection = new OleDbConnection(connectionString);
                connection.Open();
                clientName = clientComboBox.Text;
                //In the event of a client that contains a ' in it's name, it adds an additional ' so that the
                // SQL query generated here is valid
                if (clientName.Contains("'"))
                {
                    clientName = clientName.Replace("'", "''");
                }
                OleDbCommand getAccountID = new OleDbCommand("select AccountID from GTG_MSCRM.dbo.AccountBase where Name = '" + clientName + "'", connection);
                OleDbDataReader reader = getAccountID.ExecuteReader();
                while (reader.Read())
                {
                    selectedClient = reader.GetValue(0).ToString();
                }
                connection.Close();
                UpdatePOC();
                UpdateAssets();
                // Resets the poc email & phone text since the client was changed and the previous poc is not longer valid.
                poc_email.Text = "";
                poc_phone.Text = "";
            }
            catch(Exception getAccountID_ex)
            {
                MessageBox.Show(getAccountID_ex.ToString(), "Error Generating AccountID");
            }
        }

        /*UpdatePOCData()
        * The following function updates the client's POC combo box based on the client selected by the user
        */
        private void UpdatePOCData()
        {
            try
            {
                OleDbCommand getPOCEmail = new OleDbCommand("select Case WHEN EMailAddress1 is null THEN 'No Email on File' ELSE EmailAddress1 END as Email from GTG_MSCRM.dbo.ContactBase where ParentCustomerID = '" + selectedClient + "' and FullName = '" + clientPOC + "'", connection);
                OleDbDataReader reader = getPOCEmail.ExecuteReader();
                //Updates the text under the combo box with the POC's email for easy access to the user
                while (reader.Read())
                {
                    poc_email.Text = reader[0].ToString();
                }
                reader.Close();
                OleDbCommand getPOCPhone = new OleDbCommand("select CASE WHEN Telephone1 is null THEN 'No Phone Number on File' ELSE Telephone1 END as Phone from GTG_MSCRM.dbo.ContactBase where ParentCustomerID = '" + selectedClient + "' and FullName = '" + clientPOC + "'", connection);
                OleDbDataReader readerTwo = getPOCPhone.ExecuteReader();
                //Updates the text under the combo box with the POC's phone for easy access to the user
                while (readerTwo.Read())
                {
                    poc_phone.Text = readerTwo[0].ToString();
                }
                readerTwo.Close();


            }
            catch (Exception pocComboEX)
            {
                MessageBox.Show(pocComboEX.ToString(), "Error updating contact data");
            }
        }

        /*pocComboBox_SelectedIndexChanged()
        * Calls to update the POC data based on the POC selected by the user
        */
        private void pocComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            clientPOC = pocComboBox.Text;
            UpdatePOCData();
        }

        //Sets the clientArcVersion field based on the selected text in the combo box
        private void arcGISComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            clientArcVersion = arcGISComboBox.Text;
        }

        //Sets the clientStatus field based on the selected text in the combo box
        private void statusComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            clientStatus = statusComboBox.Text;
        }

        //Text field that the user can type out any issues reported during the tune up call.
        private void issuesTextBox_TextChanged(object sender, EventArgs e)
        {
            clientIssues = issuesTextBox.Text;
        }

        //Text field that the user can type out any opportunities discovered during the tune up call.
        private void oppTextBox_TextChanged(object sender, EventArgs e)
        {
            clientOpportunites = oppTextBox.Text;
        }

        //Text field that the user can type out their summary of the client during the tune up call.
        private void summaryTextBox_TextChanged(object sender, EventArgs e)
        {
            clientSummary = summaryTextBox.Text;
        }

        //Allows the enter key to select an item from a combo box
        private void enterKey_Button(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                enterKey.PerformClick();
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        }

        //Generates a .pdf based on the items entered and selected and then emails the report to the 
        //Technical Support Manager
        private void reportGenButton_Click(object sender, EventArgs e)
        {
            createReport();
            emailReport();
        }

        /*createReport()
        * Generates a .pdf document of the report based on the items selected. It requires that all fields have an item
        * selected or text entered. Once completed, it'll place the document on the user's Desktop. The following does
        * require the iTextSharp extension.
        */
        private void createReport()
        {
            try
            {
                //The following makes sure that all items are entered.
                if (clientName == null)
                {
                    MessageBox.Show("Client Name Cannot Be Empty", "Empty");
                }
                else if (clientStatus == null)
                {
                    MessageBox.Show("Client Status Cannot Be Empty", "Empty");
                }
                else if (clientPOC == null)
                {
                    MessageBox.Show("Client POC Cannot Be Empty", "Empty");
                }
                else if (clientArcVersion == null)
                {
                    MessageBox.Show("Client Arc Version Cannot Be Empty", "Empty");
                }
                else if (issuesTextBox.Text == "")
                {
                    MessageBox.Show("Open Issues Cannot Be Empty", "Empty");
                }
                else if (oppTextBox.Text == "")
                {
                    MessageBox.Show("Opportunities Cannot Be Empty", "Empty");
                }
                else if (summaryTextBox.Text == "")
                {
                    MessageBox.Show("Summary Cannot Be Empty", "Empty");
                }
                else
                {
                    //CSets the file path to the user's desktop
                    var desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    filename = Path.Combine(desktopFolder, reportAuthor + "_" + clientName + ".pdf");
                    var fullFileName = filename;
                    FileStream fileStream = new FileStream(fullFileName, FileMode.Create, FileAccess.Write, FileShare.None);
                    Document doc = new Document();
                    PdfWriter writer = PdfWriter.GetInstance(doc, fileStream);
                    doc.Open();

                    Chunk glue = new Chunk(new VerticalPositionMark());

                    Paragraph header = new Paragraph("Client Tune Up Report - " + reportAuthor, FontFactory.GetFont(FontFactory.TIMES_BOLD, 16));
                    header.Alignment = Element.ALIGN_CENTER;
                    doc.Add(header);

                    doc.Add(Chunk.NEWLINE);

                    Phrase clientPhrase = new Phrase();
                    clientPhrase.Add(new Chunk("Client: ", FontFactory.GetFont(FontFactory.TIMES_BOLD, 14)));
                    clientPhrase.Add(new Chunk(clientName, FontFactory.GetFont(FontFactory.TIMES, 12)));
                    Paragraph clientP = new Paragraph(clientPhrase);
                    doc.Add(clientP);

                    doc.Add(Chunk.NEWLINE);

                    Phrase pocPhrase = new Phrase();
                    pocPhrase.Add(new Chunk("Client POC: ", FontFactory.GetFont(FontFactory.TIMES_BOLD, 14)));
                    pocPhrase.Add(new Chunk(clientPOC, FontFactory.GetFont(FontFactory.TIMES, 12)));
                    Paragraph pocP = new Paragraph(pocPhrase);
                    doc.Add(pocP);

                    doc.Add(Chunk.NEWLINE);

                    Phrase datePhrase = new Phrase();
                    datePhrase.Add(new Chunk("Date: ", FontFactory.GetFont(FontFactory.TIMES_BOLD, 14)));
                    datePhrase.Add(new Chunk(reportDate, FontFactory.GetFont(FontFactory.TIMES, 12)));
                    Paragraph dateP = new Paragraph(datePhrase);
                    doc.Add(dateP);

                    doc.Add(Chunk.NEWLINE);

                    Phrase softPhrase = new Phrase();
                    softPhrase.Add(new Chunk("Software: ", FontFactory.GetFont(FontFactory.TIMES_BOLD, 14)));
                    Paragraph softP = new Paragraph(softPhrase);
                    doc.Add(softP);
                    softPhrase.Clear();
                    for (int i = 0; i < assetData.RowCount - 1; i++)
                    {
                        for (int j = 0; j < assetData.ColumnCount; j++)
                        {
                            softPhrase.Add(new Chunk(assetData.Rows[i].Cells[j].Value.ToString(), FontFactory.GetFont(FontFactory.TIMES, 12)));
                            softPhrase.Add(new Chunk("    "));
                        }
                        softP = new Paragraph(softPhrase);
                        doc.Add(softP);
                        softPhrase.Clear();
                    }

                    doc.Add(Chunk.NEWLINE);

                    Phrase gisPhrase = new Phrase();
                    gisPhrase.Add(new Chunk("GIS Platform: ", FontFactory.GetFont(FontFactory.TIMES_BOLD, 14)));
                    gisPhrase.Add(new Chunk(clientArcVersion, FontFactory.GetFont(FontFactory.TIMES, 12)));
                    Paragraph gisP = new Paragraph(gisPhrase);
                    doc.Add(gisP);

                    doc.Add(Chunk.NEWLINE);

                    Phrase issuesPhrase = new Phrase();
                    issuesPhrase.Add(new Chunk("Open Issues: ", FontFactory.GetFont(FontFactory.TIMES_BOLD, 14)));
                    Paragraph issuesP = new Paragraph(issuesPhrase);
                    doc.Add(issuesP);
                    issuesPhrase.Clear();
                    issuesPhrase.Add(new Chunk(clientIssues, FontFactory.GetFont(FontFactory.TIMES, 12)));
                    issuesP = new Paragraph(issuesPhrase);
                    doc.Add(issuesP);

                    doc.Add(Chunk.NEWLINE);

                    Phrase oppPhrase = new Phrase();
                    oppPhrase.Add(new Chunk("Opportunities: ", FontFactory.GetFont(FontFactory.TIMES_BOLD, 14)));
                    Paragraph oppP = new Paragraph(oppPhrase);
                    doc.Add(oppP);
                    oppPhrase.Clear();
                    oppPhrase.Add(new Chunk(clientOpportunites, FontFactory.GetFont(FontFactory.TIMES, 12)));
                    oppP = new Paragraph(oppPhrase);
                    doc.Add(oppP);

                    doc.Add(Chunk.NEWLINE);

                    Phrase sumPhrase = new Phrase();
                    sumPhrase.Add(new Chunk("Summary: " + "  (" + clientStatus + ")", FontFactory.GetFont(FontFactory.TIMES_BOLD, 14)));
                    Paragraph sumP = new Paragraph(sumPhrase);
                    doc.Add(sumP);
                    sumPhrase.Clear();
                    sumPhrase.Add(new Chunk(clientSummary, FontFactory.GetFont(FontFactory.TIMES, 12)));
                    sumP = new Paragraph(sumPhrase);
                    doc.Add(sumP);

                    doc.Add(Chunk.NEWLINE);

                    doc.Close();
                }
            }
            catch (Exception reportGen_ex)
            {
                MessageBox.Show(reportGen_ex.ToString(), "Error Generating Report");
            }
        }

        /*emailReport()
        * After the report is generated, the following function generates an email to the technical support manager
        * and send the report to them.
        */
        private void emailReport()
        {
            try
            {
                //Change recipient in the event I leave.
                emailRecipient = "bbutts@geotg.com";
                //Create the Outlook application
                Outlook.Application mailApplication = new Outlook.Application();
                //Create a new mail item
                Outlook.MailItem mailMessage = (Outlook.MailItem)mailApplication.CreateItem(Outlook.OlItemType.olMailItem);
                //Set HTMLBody
                mailMessage.HTMLBody = "Attached is the Tune Up Report for [" + clientName + "] completed by " + reportAuthor + ".";
                //Add an attachment
                String sDisplayName ="Tune Up Report";
                int iPosition = (int)mailMessage.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //now attached the file
                Outlook.Attachment mailAttachments = mailMessage.Attachments.Add(@filename, iAttachType, iPosition, sDisplayName);
                //Subject line
                mailMessage.Subject = "Tune Up Report for " + clientName;
                Outlook.Recipients mailRecipients = (Outlook.Recipients)mailMessage.Recipients;
                //Change the recipient in the next line if necessary
                Outlook.Recipient mailRecipient = (Outlook.Recipient)mailRecipients.Add(emailRecipient);
                mailRecipient.Resolve();
                //Send
                mailMessage.Send();
                //Clean up
                mailRecipient = null;
                mailRecipients = null;
                mailMessage = null;
                mailApplication = null;
                MessageBox.Show("Successfully Created Report at " + filename + "! and emailed to " + emailRecipient);
            }
            catch (Exception mail_ex)
            {
                MessageBox.Show(mail_ex.ToString(), "Error Sending Email");
            }
        }
    }
}

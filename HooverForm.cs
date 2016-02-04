using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;

namespace HooverProject
{
    public partial class HooverForm : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void button1_Click(object sender, EventArgs e)
        {

            String connectionString = "Password=S3rvice1;Persist Security Info=True;User ID=sa;Initial Catalog=Hoover;Data Source=JOSHUA\\SQLEXPRESS;Connection Timeout=5;";
            using (SqlConnection connection = new SqlConnection())
            {
                connection.ConnectionString = connectionString;
                try
                {
                    connection.Open();
                    DataTable myTable = CreateTable(connection);
                    myTable = UpdateTable(myTable);
                    createSQL(myTable);
                    connection.Close();                    
                }
                catch (Exception error)
                {
                    textBox1.Text = error.ToString();
                }
            }

        }

        public void createSQL(DataTable myTable)
        {
            String connectionString = "Password=S3rvice1;Persist Security Info=True;User ID=sa;Initial Catalog=HooverTwo;Data Source=JOSHUA\\SQLEXPRESS";
            using (SqlConnection connectionTwo = new SqlConnection())
            {
                connectionTwo.ConnectionString = connectionString;
                try
                {
                    connectionTwo.Open();
                    SqlCommand clear = new SqlCommand("DELETE FROM HooverTwo.dbo.HooverTwo", connectionTwo);
                    SqlDataReader dr = clear.ExecuteReader();
                    dr.Close();
                    foreach (DataRow row in myTable.Rows)
                    {
                        string Incident_ID = row["Incident_ID"].ToString();
                        Convert.ToInt32(Incident_ID);
                        string format = "yyyy-MM-dd HH:mm:ss";
                        string Rec_Time = row["Rec_Time"].ToString();
                        Rec_Time = Convert.ToDateTime(Rec_Time).ToString(format);
                        string Call_Type = row["Call_Type"].ToString();
                        string Address = row["Address"].ToString();
                        string Call_Status = row["Call_Status"].ToString();
                        string Unit_ID = row["Unit_ID"].ToString();
                        SqlCommand command = new SqlCommand("SELECT * FROM HooverTwo.dbo.HooverTwo WHERE Incident_ID = '" + Incident_ID + "'", connectionTwo);
                        SqlDataReader reader = command.ExecuteReader();

                        SqlCommand cmd = new SqlCommand();
                        if (reader.HasRows == true)
                        {
                            textBox1.Text = textBox1.Text + "True ";
                            cmd = new SqlCommand("UPDATE HooverTwo.dbo.HooverTwo SET Unit_ID = Unit_ID + ' ' + '" + Unit_ID + "' WHERE Incident_ID = '" + Incident_ID + "'", connectionTwo);
                        }
                        else
                        {
                            cmd = new SqlCommand("INSERT INTO HooverTwo.dbo.HooverTwo VALUES ('" + Incident_ID + "','" + Rec_Time + "','" + Call_Type + "','" + Address + "','" + Call_Status + "','" + Unit_ID + "')", connectionTwo);
 
                        }
                        reader.Close();
                        dr = cmd.ExecuteReader();
                        dr.Close();
                    }
                    connectionTwo.Close();
                }
                catch (Exception error)
                {
                    textBox1.Text = error.ToString();
                }
            }
        }

        public DataTable CreateTable(SqlConnection connection)
        {
            DataTable table = new DataTable();
            SqlCommand cmd = new SqlCommand("SELECT * FROM Hoover.dbo.Hoover", connection);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(table);
            da.Dispose();
            return table;
        }

        public DataTable UpdateTable(DataTable myTable)
        {
            DataTable table = new DataTable();
            DataTable tableTwo = new DataTable();
            DataTable updatedTable = new DataTable();
            table = myTable;
            tableTwo = myTable;
            updatedTable = myTable.Clone();
            updatedTable.Rows.Clear();

            foreach (DataRow i in table.Rows)
            {
                foreach (DataRow j in tableTwo.Rows)
                {
                    if(i["Incident_ID"].Equals(j["Incident_ID"]))
                    {
                        updatedTable.Rows.Add(i.ItemArray);
                    }
                }
            }

            return myTable;
        }
    }
}

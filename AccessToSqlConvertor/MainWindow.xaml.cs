using System;
using System.IO;
using FastMember;
using System.Linq;
using System.Data;
using System.Windows;
using Microsoft.Win32;
using System.Data.OleDb;
using System.Data.SqlClient;
using Path = System.IO.Path;
using System.Collections.Generic;
using CommonWindows;

namespace AccessToSqlConvertor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow() => InitializeComponent();
        
        public string AccessLocation { get; set; }
        public string SqlLocation { get; set; } = null;
        private void OpenAccessFile(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new()
            {
                DefaultExt = ".mdb",
                Filter = "Access File (*.mdb)|*.mdb",
                Title = "Choose access file [MDB file]",
                Multiselect = false
            };
            bool? result = openFile.ShowDialog();
            if (result is not null)
            {
                string fileName = openFile.FileName;
                AccessLocation = Access.Text = fileName;
            }
        }

        private void OpenSqlFile(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new()
            {
                DefaultExt = ".mdf",
                Filter = "SQL File (*.mdf)|*.mdf",
                Title = "Choose local dataset file [MDF file]"
            };
            bool? result = openFile.ShowDialog();
            if (result is not null)
            {
                string fileName = openFile.FileName;
                SqlLocation = Sql.Text = fileName;
            }
        }

        private void Convert(object sender, RoutedEventArgs e)
        {
            string emptySqlConnection = $"Data Source=(LocalDB)\\MSSQLLocalDB;Initial Catalog=master; Integrated Security=true;";
            string accessConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessLocation};Persist Security Info=False;";

            // Fill the List of DataTable with all the tables inside the access file
            using OleDbConnection oleDbConnection = new(accessConnectionString);
            oleDbConnection.Open();
            List<string> accessTablesName = GetAccessTables(oleDbConnection);
            List<DataTable> accessTables = new();
            foreach (string accessTableName in accessTablesName)
            {
                string accessQuery = $"SELECT * FROM [{accessTableName}]";
                try
                {
                    using OleDbDataAdapter oleDbDataAdapter = new(accessQuery, oleDbConnection);
                    DataTable accessTable = new();
                    oleDbDataAdapter.Fill(accessTable);
                    accessTables.Add(accessTable);
                }
                catch (Exception ex)
                {
                    new InfoBox("אישור", ex.Message, MessageLevel.Warning).ShowDialog();
                }
            }

            string databaseName = Path.GetFileNameWithoutExtension(AccessLocation);
            string sqlNewDatabase = SqlLocation ?? AccessLocation.Replace("mdb", "mdf");

            // If the the new MDF file is not exists, create new
            if (!File.Exists(sqlNewDatabase))
            {
                using SqlConnection sqlConnection = new(emptySqlConnection);
                sqlConnection.Open();
                try
                {
                    //foreach (string table in accessTablesName)
                    //{
                    //}
                    using SqlCommand command = sqlConnection.CreateCommand();

                    command.CommandText = string.Format("CREATE DATABASE [{0}] ON PRIMARY (NAME=[{0}], FILENAME='{1}')", databaseName, sqlNewDatabase);
                    command.ExecuteNonQuery();

                    command.CommandText = $"EXEC sp_detach_db '{databaseName}', 'true'";
                    command.ExecuteNonQuery();

                    new InfoBox("אישור", $"מסד הנתונים {databaseName} נוצר בהצלחה בנתיב:{sqlNewDatabase}", MessageLevel.OK).ShowDialog();
                }
                catch (Exception ex)
                {
                    new InfoBox("אישור", $"מסד הנתונים {databaseName}עקב שגיאה {ex.Message} לא נוצר", MessageLevel.Error).ShowDialog();
                    Environment.Exit(1);
                }
            }

            // MDF connection string
            string sqlConnectionString = $"Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename={sqlNewDatabase};Integrated Security=True;MultipleActiveResultSets=True";

            // Fill the MDF file with the tables from the access file
            using (SqlConnection sqlConnection = new(sqlConnectionString))
            {
                sqlConnection.Open();
                for (int i = 0; i < accessTables.Count; i++)
                {
                    string[] columnsNames = GetAccessColumnsNames(accessConnectionString, accessTablesName[i]);
                    string[] columnsDataType = GetAccessColumnsDataType(accessConnectionString, accessTablesName[i]);

                    string[] columns = new string[columnsNames.Length];
                    for (int j = 0; j < columnsNames.Length; j++)
                    {
                        columns[j] = $"[{columnsNames[j]}] {GetType(columnsDataType[j])}";
                    }



                    string query = $"CREATE TABLE [{accessTablesName[i]}] " +
                                   $"({string.Join(", ", columns)})";
                    try
                    {
                        using SqlCommand command = new(query, sqlConnection);
                        command.ExecuteNonQuery();
                    }
                    catch (Exception ex){ }

                    using SqlTransaction sqlTransaction = sqlConnection.BeginTransaction();
                    try
                    {
                        using SqlBulkCopy sqlBulkCopy = new(sqlConnection, SqlBulkCopyOptions.Default, sqlTransaction);
                        sqlBulkCopy.DestinationTableName = accessTablesName[i];
                        if (columnsNames is not null || columnsNames.Length > 1)
                        {
                            using ObjectReader objectReader = ObjectReader.Create(accessTables, columnsNames);
                            sqlBulkCopy.DestinationTableName = accessTablesName[i];
                            foreach (DataColumn column in accessTables[i].Columns)
                            {
                                sqlBulkCopy.ColumnMappings.Add(column.ToString(), column.ToString());
                            }
                            sqlBulkCopy.WriteToServer(accessTables[i]);
                            sqlTransaction.Commit();
                        }
                        else
                        {
                            new InfoBox("אישור", "לא נמצאו עמודות במסד הנתונים", MessageLevel.Warning).ShowDialog();
                        }
                    }
                    catch (Exception xe)
                    {
                        sqlTransaction.Rollback();
                        new InfoBox("אישור", xe.Message, MessageLevel.Warning).ShowDialog();
                    }
                }
            }
            new InfoBox("אישור", "מסד הנתונים הישן הועתק לחדש בהצלחה!", MessageLevel.OK).ShowDialog();
            WindowState = WindowState.Normal;
            Focus();
        }

        /// <summary>
        /// Get <see cref="List{DataTable}"/> of all the tables in the access database
        /// </summary>
        /// <param name="connString">The specified string to the access database</param>
        /// <returns><see cref="List{DataTable}"/> of all the tables in the access database</returns>
        private List<DataTable> GetAccessTables(string connString)
        {
            using OleDbConnection oleDbConnection = new(connString);
            oleDbConnection.Open();
            List<string> accessTablesName = GetAccessTables(oleDbConnection);
            List<DataTable> accessTables = new();
            foreach (string accessTableName in accessTablesName)
            {
                string accessQuery = $"SELECT * FROM [{accessTableName}]";
                try
                {
                    using OleDbDataAdapter oleDbDataAdapter = new(accessQuery, oleDbConnection);
                    DataTable accessTable = new();
                    oleDbDataAdapter.Fill(accessTable);
                    accessTables.Add(accessTable);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return accessTables;
        }

        /// <summary>
        /// Get all the tables name in the Access|mdb file
        /// </summary>
        /// <param name="conn">The specified open connection</param>
        /// <returns>List of all the tables name</returns>
        private List<string> GetAccessTables(OleDbConnection conn) 
            => conn.GetSchema("Tables").Select("TABLE_TYPE = 'TABLE'").Select(row => row["TABLE_NAME"].ToString()).ToList();
        
        /// <summary>
        /// Get all the tables name in the SQL|mdf file
        /// </summary>
        /// <param name="conn">The specified open connection</param>
        /// <returns>List of all the tables name</returns>
        private List<string> GetSqlTables(SqlConnection conn) 
            => conn.GetSchema("Tables").Select("TABLE_TYPE = 'TABLE'").Select(row => row["TABLE_NAME"].ToString()).ToList();
        
        /// <summary>
        /// Get all the specified table columns names
        /// </summary>
        /// <param name="tableName">Table name</param>
        /// <returns>String array of the table columns names</returns>
        private string[] GetSqlColumnsNames(SqlConnection sqlConnection, string tableName)
        {
            List<string> columnsNames = null;
            try
            {
                string query = $"SELECT column_name " +
                               $"FROM information_schema.columns " +
                               $"WHERE table_name = '{tableName}'";
                using SqlCommand command = new(query, sqlConnection);
                using SqlDataReader reader = command.ExecuteReader();
                columnsNames = new List<string>();
                while (reader.Read())
                {
                    string header = reader.GetString(0)[0].ToString().ToUpper() + reader.GetString(0).Substring(1);
                    columnsNames.Add(header);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return columnsNames?.ToArray();
        }

        /// <summary>
        /// Get all the specified table columns names
        /// </summary>
        /// <param name="tableName">Table name</param>
        /// <returns>String array of the table columns names</returns>
        private string[] GetAccessColumnsNames(string connString, string tableName)
        {
            List<string> columnsNames = new();
            using OleDbConnection oleDbConnection = new(connString);
            oleDbConnection.Open();
            try
            {
                string query = $"SELECT TOP 1 * FROM [{tableName}]";
                using OleDbCommand command = new(query, oleDbConnection);
                using OleDbDataReader reader = command.ExecuteReader(CommandBehavior.SchemaOnly);
                DataTable data = reader.GetSchemaTable();
                columnsNames.AddRange(from DataRow row in data.Rows select row.Field<string>(0));
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return columnsNames?.ToArray();
        }

        /// <summary>
        /// Get all the specified table columns data types
        /// </summary>
        /// <param name="tableName">Table name</param>
        /// <returns>String array of the table columns data type</returns>
        private string[] GetAccessColumnsDataType(string connString, string tableName)
        {
            List<string> columnsNames = new();
            using OleDbConnection oleDbConnection = new(connString);
            oleDbConnection.Open();
            try
            {
                string query = $"SELECT TOP 1 * FROM [{tableName}]";
                using OleDbCommand command = new(query, oleDbConnection);
                using OleDbDataReader reader = command.ExecuteReader(CommandBehavior.SchemaOnly);
                //DataTable data = oleDbConnection.GetSchema("Columns");
                DataTable data = reader.GetSchemaTable();
                columnsNames.AddRange(from DataRow row in data.Rows select row.Field<Type>(5).FullName.ToString());
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return columnsNames?.ToArray();
        }

        /// <summary>
        /// Convert from <see cref="Type"/> to <see cref="string"/> that match SQL Type
        /// </summary>
        /// <param name="type">The type to be converted</param>
        /// <returns>The converted type</returns>
        private string GetType(string type)
            => type.ToLower() switch
            {
                "system.int32" => "INT",
                "system.long" => "BIGINT",
                "system.single" => "REAL",
                "system.float" => "FLOAT",
                "system.bool" => "BOOLEAN",
                "system.datetime" => "DATE",
                "system.char" => "CHAR(255)",
                "system.string" => "NVARCHAR(MAX)",
                _ => "NVARCHAR(255)"
            };
    }
}

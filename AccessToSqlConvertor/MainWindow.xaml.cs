using System;
using System.IO;
using System.Linq;
using System.Data;
  /*  Custom DLL for InfoBox  */
 /**/ using CommonWindows; /**/
/*~ ~ ~ ~ ~ ~ ~~ ~ ~ ~ ~ ~ ~*/
using System.Windows;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Data.SqlClient;
using Path = System.IO.Path;
using System.Collections.Generic;
using Result = System.Windows.Forms.DialogResult;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace AccessToSqlConvertor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Constructor

        public MainWindow() => InitializeComponent();

        #endregion

        #region Properties
        public string AccessLocation { get; set; }
        public string SqlLocation { get; set; } = null;

        #endregion

        #region Buttons click event

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

            // Choose to save the new .MDF file in the same folder or in other folder
            string folderName = null;
            YesNoWindow yesNoWindow = new("האם ברצונך לבחור תיקית יעד שונה לשמירת מסד הנתונים?", "כן", "לא");
            yesNoWindow.ShowDialog();
            if (yesNoWindow.ResultYes)
            {
                using var folderBrowser = new FolderBrowserDialog();
                Result result = folderBrowser.ShowDialog();
                if (result == Result.OK && !string.IsNullOrWhiteSpace(folderBrowser.SelectedPath))
                {
                    folderName = folderBrowser.SelectedPath;
                }
            }

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

            // Data base name
            string databaseName = Path.GetFileNameWithoutExtension(AccessLocation);

            // New .MDF file save location
            string sqlNewDatabase = folderName is null ? SqlLocation ?? AccessLocation.Replace("mdb", "mdf") : $"{folderName}\\{databaseName}.mdf";

            // If the the new MDF file is not exists, create new
            if (!File.Exists(sqlNewDatabase))
            {
                using SqlConnection sqlConnection = new(emptySqlConnection);
                sqlConnection.Open();
                try
                {
                    using SqlCommand command = sqlConnection.CreateCommand();

                    command.CommandText = string.Format("CREATE DATABASE [{0}] ON PRIMARY (NAME=[{0}], FILENAME='{1}')", databaseName, sqlNewDatabase);
                    command.ExecuteNonQuery();

                    command.CommandText = $"EXEC sp_detach_db '{databaseName}', 'true'";
                    command.ExecuteNonQuery();

                    new InfoBox("אישור", $"מסד הנתונים {databaseName} נוצר בהצלחה בנתיב:\n{sqlNewDatabase}", MessageLevel.OK).ShowDialog();
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
                    string query = $"CREATE TABLE [{accessTablesName[i]}] ({string.Join(", ", columns)})";
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
                            //using ObjectReader objectReader = ObjectReader.Create(accessTables, columnsNames);
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

        #endregion

        #region Methods

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
                    new InfoBox("אישור", ex.Message, MessageLevel.Warning).ShowDialog();
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
                DataTable data = reader.GetSchemaTable();
                columnsNames = new List<string>();
                columnsNames.AddRange(from DataRow row in data.Rows select row.Field<string>(0));
            }
            catch (Exception e)
            {
                new InfoBox("אישור", e.Message, MessageLevel.Warning).ShowDialog();
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
                new InfoBox("אישור", e.Message, MessageLevel.Warning).ShowDialog();
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
                new InfoBox("אישור", e.Message, MessageLevel.Warning).ShowDialog();
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
                "system.long"      => "BIGINT",
                "system.boolean" => "BIT",
                "system.char"      => "CHAR(50)",
                "system.int32"     => "INT",
                "system.string"    => "NVARCHAR(255)",
                "system.decimal"   => "decimal",
                "system.datetime"  => "DATE",
                "system.float"
                or "system.double"
                or "system.single" => "FLOAT",
                _ => "NVARCHAR(255)"
            };

        #endregion
    }
}

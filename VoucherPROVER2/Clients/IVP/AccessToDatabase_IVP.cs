using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VoucherPROVER2.Clients.IVP
{
    public class AccessToDatabase_IVP
    {
        readonly string client = GlobalVariables.client;
        public static string GetAccessConnectionString()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string fileName = "CheckDatabase.accdb";
            string resourcePath = Path.Combine(baseDirectory, fileName);
            string accessConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={resourcePath};Persist Security Info=False;";
            return accessConnectionString;
        }

        public static string GetQBConnectionString()
        {
            string qbConnectionString = "DSN=QuickBooks Data;";
            return qbConnectionString;
        }

        //List<string> tableNamesWithItemReceipt = new List<string> { "Account", "Bill", "BillExpenseLine", "BillItemLine", "BillPaymentCheck", "BillPaymentCheckLine", "Check", "CheckExpenseLine", "CheckItemLine", "Company", "Item", "ItemReceipt", "ItemReceiptExpenseLine", "ItemReceiptItemLine",  "Transaction", "Vendor" };
        List<string> tableNamesWithItemReceipt = new List<string> { "Account","Bill", "BillExpenseLine", "BillItemLine", "BillPaymentCheck", "BillPaymentCheckLine","Check", "CheckExpenseLine", "CheckItemLine", "Company", "Item", "ItemReceipt", "ItemReceiptExpenseLine", "ItemReceiptItemLine",
            "Vendor" };

        List<string> tableNamesWithoutItemReceipt = new List<string> { "Account",
            "Bill", "BillExpenseLine", "BillItemLine", "BillPaymentCheck", "BillPaymentCheckLine",
            "Check", "CheckExpenseLine", "CheckItemLine", "Company", "Item", "Vendor" };

        public async Task FetchAndSaveData()
        {
            string odbcConnectionString = GetQBConnectionString();
            string accessConnectionString = GetAccessConnectionString();

            List<string> tableNames;

            if (GlobalVariables.includeItemReceipt)
            {
                tableNames = tableNamesWithItemReceipt;
            }
            else
            {
                tableNames = tableNamesWithoutItemReceipt;
            }

            using (OdbcConnection odbcConnection = new OdbcConnection(odbcConnectionString))
            {
                try
                {
                    await odbcConnection.OpenAsync();

                    using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
                    {
                        await accessConnection.OpenAsync();

                        foreach (var tableName in tableNames)
                        {
                            // Ensure table name is enclosed in square brackets
                            string formattedTableName = $"[{tableName}]";

                            // Create ODBC command to select data from the QuickBooks Data DSN
                            OdbcCommand odbcCommand = odbcConnection.CreateCommand();
                            odbcCommand.CommandText = $"SELECT * FROM {tableName}";

                            // Execute the ODBC command to fetch the data
                            using (OdbcDataReader reader = (OdbcDataReader)await odbcCommand.ExecuteReaderAsync())
                            {
                                // Create the destination table name with "QB_" prefix
                                //string destinationTableName = $"QB_{tableName}";
                                //string destinationTableName = $"{tableName}";
                                string destinationTableName = formattedTableName;

                                // Construct the Access SQL command for inserting data
                                OleDbCommand accessCommand = accessConnection.CreateCommand();
                                string columnNames = GetColumnNames(reader);
                                string parameterNames = GetParameterNames(reader);
                                accessCommand.CommandText = $"INSERT INTO {destinationTableName} ({columnNames}) VALUES ({parameterNames})";

                                // Add parameters to the Access command
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    accessCommand.Parameters.Add(new OleDbParameter($"@param{i}", reader.GetFieldType(i)));
                                }

                                // Transfer data row by row
                                while (await reader.ReadAsync())
                                {
                                    for (int i = 0; i < reader.FieldCount; i++)
                                    {
                                        accessCommand.Parameters[$"@param{i}"].Value = reader.GetValue(i);
                                    }

                                    try
                                    {
                                        await accessCommand.ExecuteNonQueryAsync();
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"An error occurred while inserting data into {destinationTableName}: {ex.Message}\nSQL: {accessCommand.CommandText}");
                                        return;
                                    }
                                }
                            }
                        }
                        accessConnection.Close();
                    }
                    odbcConnection.Close();

                    MessageBox.Show("Data transfer completed successfully.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while fetching data: " + ex.Message);
                }
            }
        }

        // Helper method to get column names from the OdbcDataReader
        private string GetColumnNames(OdbcDataReader reader)
        {
            List<string> columnNames = new List<string>();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                columnNames.Add($"[{reader.GetName(i)}]");
            }
            return string.Join(", ", columnNames);
        }

        // Helper method to get parameter names for the SQL command
        private string GetParameterNames(OdbcDataReader reader)
        {
            List<string> parameterNames = new List<string>();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                parameterNames.Add($"@param{i}");
            }
            return string.Join(", ", parameterNames);
        }

        public void DeleteSpecifiedTablesData()
        {
            string accessConnectionString = GetAccessConnectionString();

            List<string> tableNames;

            if (GlobalVariables.includeItemReceipt)
            {
                tableNames = tableNamesWithItemReceipt;
            }
            else
            {
                tableNames = tableNamesWithoutItemReceipt;
            }

            using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
            {
                accessConnection.Open();

                foreach (var tableName in tableNames)
                {
                    // Construct the SQL command to delete all data from the table
                    OleDbCommand accessCommand = accessConnection.CreateCommand();
                    accessCommand.CommandText = $"DELETE FROM [{tableName}]";

                    try
                    {
                        accessCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"An error occurred while deleting data from {tableName}: {ex.Message}");
                        return;
                    }
                }
                accessConnection.Close();
            }

            Console.WriteLine("All data from specified tables has been deleted.");
        }

        public void SaveSignatoryData(int choice, string name, string position)
        {
            string accessConnectionString = GetAccessConnectionString();

            try
            {
                using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
                {
                    connection.Open();

                    string selectQuery = "SELECT COUNT(*) FROM Signatory";
                    int rowCount;

                    using (OleDbCommand selectCommand = new OleDbCommand(selectQuery, connection))
                    {
                        rowCount = (int)selectCommand.ExecuteScalar();
                    }

                    string signatoryQuery = null;

                    if (rowCount > 0)
                    {
                        if (client == "LEADS")
                        {
                            switch (choice)
                            {
                                case 1:
                                    signatoryQuery = "UPDATE Signatory SET PreparedByName = ?, PreparedByPosition = ?";
                                    break;
                                case 2:
                                    signatoryQuery = "UPDATE Signatory SET ReviewedByName = ?, ReviewedByPosition = ?";
                                    break;
                                /*case 3:
                                    signatoryQuery = "UPDATE Signatory SET RecommendingApprovalName = ?, RecommendingApprovalPosition = ?";
                                    break;*/
                                case 3:
                                    signatoryQuery = "UPDATE Signatory SET ApprovedByName = ?, ApprovedByPosition = ?";
                                    break;
                                case 4:
                                    signatoryQuery = "UPDATE Signatory SET ReceivedByName = ?, ReceivedByPosition = ?";
                                    break;
                                default:
                                    break;
                            }
                        }
                        else if (client == "KAYAK")
                        {
                            switch (choice)
                            {
                                case 1:
                                    signatoryQuery = "UPDATE Signatory SET PreparedByName = ?, PreparedByPosition = ?";
                                    break;
                                case 2:
                                    signatoryQuery = "UPDATE Signatory SET ReviewedByName = ?, ReviewedByPosition = ?";
                                    break;
                                /*case 3:
                                    signatoryQuery = "UPDATE Signatory SET RecommendingApprovalName = ?, RecommendingApprovalPosition = ?";
                                    break;*/
                                case 3:
                                    signatoryQuery = "UPDATE Signatory SET ApprovedByName = ?, ApprovedByPosition = ?";
                                    break;
                                case 4:
                                    signatoryQuery = "UPDATE Signatory SET ReceivedByName = ?, ReceivedByPosition = ?";
                                    break;
                                default:
                                    break;
                            }
                        }
                        else if (client == "CPI")
                        {
                            switch (choice)
                            {
                                case 1:
                                    signatoryQuery = "UPDATE Signatory SET PreparedByName = ?, PreparedByPosition = ?";
                                    break;
                                case 2:
                                    signatoryQuery = "UPDATE Signatory SET ReviewedByName = ?, ReviewedByPosition = ?";
                                    break;
                                /*case 3:
                                    signatoryQuery = "UPDATE Signatory SET RecommendingApprovalName = ?, RecommendingApprovalPosition = ?";
                                    break;*/
                                case 3:
                                    signatoryQuery = "UPDATE Signatory SET ApprovedByName = ?, ApprovedByPosition = ?";
                                    break;
                                case 4:
                                    signatoryQuery = "UPDATE Signatory SET ReceivedByName = ?, ReceivedByPosition = ?";
                                    break;
                                default:
                                    break;
                            }
                        }
                        else if (client == "IVP")
                        {
                            switch (choice)
                            {
                                case 1:
                                    signatoryQuery = "UPDATE Signatory SET PreparedByName = ?, PreparedByPosition = ?";
                                    break;
                                case 2:
                                    signatoryQuery = "UPDATE Signatory SET ReviewedByName = ?, ReviewedByPosition = ?";
                                    break;
                                /*case 3:
                                    signatoryQuery = "UPDATE Signatory SET RecommendingApprovalName = ?, RecommendingApprovalPosition = ?";
                                    break;*/
                                case 3:
                                    signatoryQuery = "UPDATE Signatory SET ApprovedByName = ?, ApprovedByPosition = ?";
                                    break;
                                case 4:
                                    signatoryQuery = "UPDATE Signatory SET ReceivedByName = ?, ReceivedByPosition = ?";
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                    else
                    {
                        if (client == "LEADS")
                        {
                            switch (choice)
                            {
                                case 1:
                                    signatoryQuery = "INSERT INTO Signatory (PreparedByName, PreparedByPosition) VALUES (?, ?)";
                                    break;

                                case 2:
                                    signatoryQuery = "INSERT INTO Signatory (ReviewedByName, ReviewedByPosition) VALUES (?, ?)";
                                    break;

                                /*case 3:
                                    signatoryQuery = "INSERT INTO Signatory (RecommendingApprovalName, RecommendingApprovalPosition) VALUES (?, ?)";
                                    break;*/

                                case 3:
                                    signatoryQuery = "INSERT INTO Signatory (ApprovedByName, ApprovedByPosition) VALUES (?, ?)";
                                    break;

                                case 4:
                                    signatoryQuery = "INSERT INTO Signatory (ReceivedByName, ReceivedByPosition) VALUES (?, ?)";
                                    break;

                                default:
                                    break;
                            }
                        }
                        else
                        {
                            switch (choice)
                            {
                                case 1:
                                    signatoryQuery = "INSERT INTO Signatory (PreparedByName, PreparedByPosition) VALUES (?, ?)";
                                    break;

                                case 2:
                                    signatoryQuery = "INSERT INTO Signatory (ReviewedByName, ReviewedByPosition) VALUES (?, ?)";
                                    break;

                                case 3:
                                    signatoryQuery = "INSERT INTO Signatory (RecommendingApprovalName, RecommendingApprovalPosition) VALUES (?, ?)";
                                    break;

                                case 4:
                                    signatoryQuery = "INSERT INTO Signatory (ApprovedByName, ApprovedByPosition) VALUES (?, ?)";
                                    break;

                                case 5:
                                    signatoryQuery = "INSERT INTO Signatory (ReceivedByName, ReceivedByPosition) VALUES (?, ?)";
                                    break;

                                default:
                                    break;
                            }
                        }
                    }

                    using (OleDbCommand signatoryCommand = new OleDbCommand(signatoryQuery, connection))
                    {
                        if (client == "LEADS")
                        {
                            switch (choice)
                            {
                                case 1:
                                    signatoryCommand.Parameters.AddWithValue("@PreparedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@PreparedByPosition", position);
                                    break;

                                case 2:
                                    signatoryCommand.Parameters.AddWithValue("@ReviewedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ReviewedByPosition", position);
                                    break;

                                /*case 3:
                                    signatoryCommand.Parameters.AddWithValue("@RecommendingApprovalName", name);
                                    signatoryCommand.Parameters.AddWithValue("@RecommendingApprovalPosition", position);
                                    break;*/

                                case 3:
                                    signatoryCommand.Parameters.AddWithValue("@ApprovedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ApprovedByPosition", position);
                                    break;

                                case 4:
                                    signatoryCommand.Parameters.AddWithValue("@ReceivedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ReceivedByPosition", position);
                                    break;

                                default:
                                    break;
                            }
                        }
                        else if (client == "KAYAK")
                        {
                            switch (choice)
                            {
                                case 1:
                                    signatoryCommand.Parameters.AddWithValue("@PreparedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@PreparedByPosition", position);
                                    break;

                                case 2:
                                    signatoryCommand.Parameters.AddWithValue("@ReviewedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ReviewedByPosition", position);
                                    break;

                                /*case 3:
                                    signatoryCommand.Parameters.AddWithValue("@RecommendingApprovalName", name);
                                    signatoryCommand.Parameters.AddWithValue("@RecommendingApprovalPosition", position);
                                    break;*/

                                case 3:
                                    signatoryCommand.Parameters.AddWithValue("@ApprovedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ApprovedByPosition", position);
                                    break;

                                case 4:
                                    signatoryCommand.Parameters.AddWithValue("@ReceivedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ReceivedByPosition", position);
                                    break;

                                default:
                                    break;
                            }
                        }
                        else if (client == "CPI")
                        {
                            switch (choice)
                            {
                                case 1:
                                    signatoryCommand.Parameters.AddWithValue("@PreparedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@PreparedByPosition", position);
                                    break;

                                case 2:
                                    signatoryCommand.Parameters.AddWithValue("@ReviewedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ReviewedByPosition", position);
                                    break;

                                /*case 3:
                                    signatoryCommand.Parameters.AddWithValue("@RecommendingApprovalName", name);
                                    signatoryCommand.Parameters.AddWithValue("@RecommendingApprovalPosition", position);
                                    break;*/

                                case 3:
                                    signatoryCommand.Parameters.AddWithValue("@ApprovedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ApprovedByPosition", position);
                                    break;

                                case 4:
                                    signatoryCommand.Parameters.AddWithValue("@ReceivedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ReceivedByPosition", position);
                                    break;

                                default:
                                    break;
                            }
                        }
                        else if (client == "IVP")
                        {
                            switch (choice)
                            {
                                case 1:
                                    signatoryCommand.Parameters.AddWithValue("@PreparedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@PreparedByPosition", position);
                                    break;

                                case 2:
                                    signatoryCommand.Parameters.AddWithValue("@ReviewedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ReviewedByPosition", position);
                                    break;

                                /*case 3:
                                    signatoryCommand.Parameters.AddWithValue("@RecommendingApprovalName", name);
                                    signatoryCommand.Parameters.AddWithValue("@RecommendingApprovalPosition", position);
                                    break;*/

                                case 3:
                                    signatoryCommand.Parameters.AddWithValue("@ApprovedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ApprovedByPosition", position);
                                    break;

                                case 4:
                                    signatoryCommand.Parameters.AddWithValue("@ReceivedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ReceivedByPosition", position);
                                    break;

                                default:
                                    break;
                            }
                        }
                        else
                        {
                            switch (choice)
                            {
                                case 1:
                                    signatoryCommand.Parameters.AddWithValue("@PreparedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@PreparedByPosition", position);
                                    break;

                                case 2:
                                    signatoryCommand.Parameters.AddWithValue("@ReviewedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ReviewedByPosition", position);
                                    break;

                                case 3:
                                    signatoryCommand.Parameters.AddWithValue("@RecommendingApprovalName", name);
                                    signatoryCommand.Parameters.AddWithValue("@RecommendingApprovalPosition", position);
                                    break;

                                case 4:
                                    signatoryCommand.Parameters.AddWithValue("@ApprovedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ApprovedByPosition", position);
                                    break;

                                case 5:
                                    signatoryCommand.Parameters.AddWithValue("@ReceivedByName", name);
                                    signatoryCommand.Parameters.AddWithValue("@ReceivedByPosition", position);
                                    break;

                                default:
                                    break;
                            }
                        }

                        int rowsAffected = signatoryCommand.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            Console.WriteLine("Data inserted/updated successfully.");
                        }
                        else
                        {
                            Console.WriteLine("No rows were affected.");
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while updating Signatory table: {ex.Message}");
            }
        }

        public (string Name, string Position) RetrieveSignatoryData(int choice)
        {
            string name = null;
            string position = null;

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
                {
                    connection.Open();

                    string query = null;

                    if (client == "LEADS")
                    {
                        switch (choice)
                        {
                            case 1:
                                query = "SELECT TOP 1 PreparedByName, PreparedByPosition FROM Signatory";
                                break;

                            case 2:
                                query = "SELECT TOP 1 ReviewedByName, ReviewedByPosition FROM Signatory";
                                break;

                            /*case 3:
                                query = "SELECT TOP 1 RecommendingApprovalName, RecommendingApprovalPosition FROM Signatory";
                                break;*/

                            case 3:
                                query = "SELECT TOP 1 ApprovedByName, ApprovedByPosition FROM Signatory";
                                break;

                            case 4:
                                query = "SELECT TOP 1 ReceivedByName, ReceivedByPosition FROM Signatory";
                                break;

                            default:
                                break;
                        }
                    }
                    else if (client == "CPI")
                    {
                        switch (choice)
                        {
                            case 1:
                                query = "SELECT TOP 1 PreparedByName, PreparedByPosition FROM Signatory";
                                break;

                            case 2:
                                query = "SELECT TOP 1 ReviewedByName, ReviewedByPosition FROM Signatory";
                                break;

                            /*case 3:
                                query = "SELECT TOP 1 RecommendingApprovalName, RecommendingApprovalPosition FROM Signatory";
                                break;*/

                            case 3:
                                query = "SELECT TOP 1 ApprovedByName, ApprovedByPosition FROM Signatory";
                                break;

                            case 4:
                                query = "SELECT TOP 1 ReceivedByName, ReceivedByPosition FROM Signatory";
                                break;

                            default:
                                break;
                        }
                    }
                    else if (client == "IVP")
                    {
                        switch (choice)
                        {
                            case 1:
                                query = "SELECT TOP 1 PreparedByName, PreparedByPosition FROM Signatory";
                                break;

                            case 2:
                                query = "SELECT TOP 1 ReviewedByName, ReviewedByPosition FROM Signatory";
                                break;

                            /*case 3:
                                query = "SELECT TOP 1 RecommendingApprovalName, RecommendingApprovalPosition FROM Signatory";
                                break;*/

                            case 3:
                                query = "SELECT TOP 1 ApprovedByName, ApprovedByPosition FROM Signatory";
                                break;

                            case 4:
                                query = "SELECT TOP 1 ReceivedByName, ReceivedByPosition FROM Signatory";
                                break;

                            default:
                                break;
                        }
                    }
                    else
                    {
                        switch (choice)
                        {
                            case 1:
                                query = "SELECT TOP 1 PreparedByName, PreparedByPosition FROM Signatory";
                                break;

                            case 2:
                                query = "SELECT TOP 1 ReviewedByName, ReviewedByPosition FROM Signatory";
                                break;

                            case 3:
                                query = "SELECT TOP 1 RecommendingApprovalName, RecommendingApprovalPosition FROM Signatory";
                                break;

                            case 4:
                                query = "SELECT TOP 1 ApprovedByName, ApprovedByPosition FROM Signatory";
                                break;

                            case 5:
                                query = "SELECT TOP 1 ReceivedByName, ReceivedByPosition FROM Signatory";
                                break;

                            default:
                                break;
                        }
                    }

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (client == "LEADS")
                                {
                                    switch (choice)
                                    {
                                        case 1:
                                            name = reader["PreparedByName"].ToString();
                                            position = reader["PreparedByPosition"].ToString();
                                            break;

                                        case 2:
                                            name = reader["ReviewedByName"].ToString();
                                            position = reader["ReviewedByPosition"].ToString();
                                            break;

                                        /*case 3:
                                            name = reader["RecommendingApprovalName"].ToString();
                                            position = reader["RecommendingApprovalPosition"].ToString();
                                            break;*/

                                        case 3:
                                            name = reader["ApprovedByName"].ToString();
                                            position = reader["ApprovedByPosition"].ToString();
                                            break;

                                        case 4:
                                            name = reader["ReceivedByName"].ToString();
                                            position = reader["ReceivedByPosition"].ToString();
                                            break;

                                        default:
                                            break;
                                    }
                                }
                                else if (client == "CPI")
                                {
                                    switch (choice)
                                    {
                                        case 1:
                                            name = reader["PreparedByName"].ToString();
                                            position = reader["PreparedByPosition"].ToString();
                                            break;

                                        case 2:
                                            name = reader["ReviewedByName"].ToString();
                                            position = reader["ReviewedByPosition"].ToString();
                                            break;

                                        /*case 3:
                                            name = reader["RecommendingApprovalName"].ToString();
                                            position = reader["RecommendingApprovalPosition"].ToString();
                                            break;*/

                                        case 3:
                                            name = reader["ApprovedByName"].ToString();
                                            position = reader["ApprovedByPosition"].ToString();
                                            break;

                                        case 4:
                                            name = reader["ReceivedByName"].ToString();
                                            position = reader["ReceivedByPosition"].ToString();
                                            break;

                                        default:
                                            break;
                                    }
                                }
                                else if (client == "IVP")
                                {
                                    switch (choice)
                                    {
                                        case 1:
                                            name = reader["PreparedByName"].ToString();
                                            position = reader["PreparedByPosition"].ToString();
                                            break;

                                        case 2:
                                            name = reader["ReviewedByName"].ToString();
                                            position = reader["ReviewedByPosition"].ToString();
                                            break;

                                        /*case 3:
                                            name = reader["RecommendingApprovalName"].ToString();
                                            position = reader["RecommendingApprovalPosition"].ToString();
                                            break;*/

                                        case 3:
                                            name = reader["ApprovedByName"].ToString();
                                            position = reader["ApprovedByPosition"].ToString();
                                            break;

                                        case 4:
                                            name = reader["ReceivedByName"].ToString();
                                            position = reader["ReceivedByPosition"].ToString();
                                            break;

                                        default:
                                            break;
                                    }
                                }
                                else
                                {
                                    switch (choice)
                                    {
                                        case 1:
                                            name = reader["PreparedByName"].ToString();
                                            position = reader["PreparedByPosition"].ToString();
                                            break;

                                        case 2:
                                            name = reader["ReviewedByName"].ToString();
                                            position = reader["ReviewedByPosition"].ToString();
                                            break;

                                        case 3:
                                            name = reader["RecommendingApprovalName"].ToString();
                                            position = reader["RecommendingApprovalPosition"].ToString();
                                            break;

                                        case 4:
                                            name = reader["ApprovedByName"].ToString();
                                            position = reader["ApprovedByPosition"].ToString();
                                            break;

                                        case 5:
                                            name = reader["ReceivedByName"].ToString();
                                            position = reader["ReceivedByPosition"].ToString();
                                            break;

                                        default:
                                            break;
                                    }
                                }
                            }
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while retrieving signatory data: {ex.Message}");
            }

            return (name, position);
        }

        public (string PreparedByName, string PreparedByPosition, string ReviewedByName, string ReviewedByPosition, string RecommendingApprovalName, string RecommendingApprovalPosition, string ApprovedByName, string ApprovedByPosition, string ReceivedByName, string ReceivedByPosition) RetrieveAllSignatoryData()
        {
            string preparedByName = null;
            string preparedByPosition = null;
            string reviewedByName = null;
            string reviewedByPosition = null;
            string recommendingApprovalName = null;
            string recommendingApprovalPosition = null;
            string approvedByName = null;
            string approvedByPosition = null;
            string receivedByName = null;
            string receivedByPosition = null;

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
                {
                    connection.Open();

                    string query;

                    if (GlobalVariables.client == "LEADS")
                    {
                        query = "SELECT TOP 1 " +
                        "PreparedByName, PreparedByPosition, " +
                        "ReviewedByName, ReviewedByPosition, " +
                        //"RecommendingApprovalName, RecommendingApprovalPosition, " +
                        "ApprovedByName, ApprovedByPosition, " +
                        "ReceivedByName, ReceivedByPosition " +
                        "FROM Signatory";
                    }
                    else if (GlobalVariables.client == "KAYAK")
                    {
                        query = "SELECT TOP 1 " +
                        "PreparedByName, PreparedByPosition, " +
                        "ReviewedByName, ReviewedByPosition, " +
                        //"RecommendingApprovalName, RecommendingApprovalPosition, " +
                        "ApprovedByName, ApprovedByPosition, " +
                        "ReceivedByName, ReceivedByPosition " +
                        "FROM Signatory";
                    }
                    else if (GlobalVariables.client == "CPI")
                    {
                        query = "SELECT TOP 1 " +
                        "PreparedByName, PreparedByPosition, " +
                        "ReviewedByName, ReviewedByPosition, " +
                        //"RecommendingApprovalName, RecommendingApprovalPosition, " +
                        "ApprovedByName, ApprovedByPosition, " +
                        "ReceivedByName, ReceivedByPosition " +
                        "FROM Signatory";
                    }
                    else if (GlobalVariables.client == "IVP")
                    {
                        query = "SELECT TOP 1 " +
                        "PreparedByName, PreparedByPosition, " +
                        "ReviewedByName, ReviewedByPosition, " +
                        //"RecommendingApprovalName, RecommendingApprovalPosition, " +
                        "ApprovedByName, ApprovedByPosition, " +
                        "ReceivedByName, ReceivedByPosition " +
                        "FROM Signatory";
                    }
                    else
                    {
                        query = "SELECT TOP 1 " +
                        "PreparedByName, PreparedByPosition, " +
                        "ReviewedByName, ReviewedByPosition, " +
                        "RecommendingApprovalName, RecommendingApprovalPosition, " +
                        "ApprovedByName, ApprovedByPosition, " +
                        "ReceivedByName, ReceivedByPosition " +
                        "FROM Signatory";
                    }

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (client == "LEADS")
                                {
                                    preparedByName = reader["PreparedByName"].ToString();
                                    preparedByPosition = reader["PreparedByPosition"].ToString();

                                    reviewedByName = reader["ReviewedByName"].ToString();
                                    reviewedByPosition = reader["ReviewedByPosition"].ToString();

                                    //recommendingApprovalName = reader["RecommendingApprovalName"].ToString();
                                    //recommendingApprovalPosition = reader["RecommendingApprovalPosition"].ToString();

                                    approvedByName = reader["ApprovedByName"].ToString();
                                    approvedByPosition = reader["ApprovedByPosition"].ToString();

                                    receivedByName = reader["ReceivedByName"].ToString();
                                    receivedByPosition = reader["ReceivedByPosition"].ToString();
                                }
                                if (client == "KAYAK")
                                {
                                    preparedByName = reader["PreparedByName"].ToString();
                                    preparedByPosition = reader["PreparedByPosition"].ToString();

                                    reviewedByName = reader["ReviewedByName"].ToString();
                                    reviewedByPosition = reader["ReviewedByPosition"].ToString();

                                    //recommendingApprovalName = reader["RecommendingApprovalName"].ToString();
                                    //recommendingApprovalPosition = reader["RecommendingApprovalPosition"].ToString();

                                    approvedByName = reader["ApprovedByName"].ToString();
                                    approvedByPosition = reader["ApprovedByPosition"].ToString();

                                    receivedByName = reader["ReceivedByName"].ToString();
                                    receivedByPosition = reader["ReceivedByPosition"].ToString();
                                }
                                if (client == "CPI")
                                {
                                    preparedByName = reader["PreparedByName"].ToString();
                                    preparedByPosition = reader["PreparedByPosition"].ToString();

                                    reviewedByName = reader["ReviewedByName"].ToString();
                                    reviewedByPosition = reader["ReviewedByPosition"].ToString();

                                    //recommendingApprovalName = reader["RecommendingApprovalName"].ToString();
                                    //recommendingApprovalPosition = reader["RecommendingApprovalPosition"].ToString();

                                    approvedByName = reader["ApprovedByName"].ToString();
                                    approvedByPosition = reader["ApprovedByPosition"].ToString();

                                    receivedByName = reader["ReceivedByName"].ToString();
                                    receivedByPosition = reader["ReceivedByPosition"].ToString();
                                }
                                if (client == "IVP")
                                {
                                    preparedByName = reader["PreparedByName"].ToString();
                                    preparedByPosition = reader["PreparedByPosition"].ToString();

                                    reviewedByName = reader["ReviewedByName"].ToString();
                                    reviewedByPosition = reader["ReviewedByPosition"].ToString();

                                    //recommendingApprovalName = reader["RecommendingApprovalName"].ToString();
                                    //recommendingApprovalPosition = reader["RecommendingApprovalPosition"].ToString();

                                    approvedByName = reader["ApprovedByName"].ToString();
                                    approvedByPosition = reader["ApprovedByPosition"].ToString();

                                    receivedByName = reader["ReceivedByName"].ToString();
                                    receivedByPosition = reader["ReceivedByPosition"].ToString();
                                }
                                else
                                {
                                    preparedByName = reader["PreparedByName"].ToString();
                                    preparedByPosition = reader["PreparedByPosition"].ToString();

                                    reviewedByName = reader["ReviewedByName"].ToString();
                                    reviewedByPosition = reader["ReviewedByPosition"].ToString();

                                    recommendingApprovalName = reader["RecommendingApprovalName"].ToString();
                                    recommendingApprovalPosition = reader["RecommendingApprovalPosition"].ToString();

                                    approvedByName = reader["ApprovedByName"].ToString();
                                    approvedByPosition = reader["ApprovedByPosition"].ToString();

                                    receivedByName = reader["ReceivedByName"].ToString();
                                    receivedByPosition = reader["ReceivedByPosition"].ToString();
                                }
                            }
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while retrieving all signatory data: {ex.Message}");
            }

            if (GlobalVariables.client == "LEADS")
            {
                recommendingApprovalName = "";
                recommendingApprovalPosition = "";
            }
            else if (GlobalVariables.client == "KAYAK")
            {
                recommendingApprovalName = "";
                recommendingApprovalPosition = "";
            }
            else if (GlobalVariables.client == "CPI")
            {
                recommendingApprovalName = "";
                recommendingApprovalPosition = "";
            }
            else if (GlobalVariables.client == "IVP")
            {
                recommendingApprovalName = "";
                recommendingApprovalPosition = "";
            }

            return (
                preparedByName, preparedByPosition,
                reviewedByName, reviewedByPosition,
                recommendingApprovalName, recommendingApprovalPosition,
                approvedByName, approvedByPosition,
                receivedByName, receivedByPosition
                );
        }


        // ----------------------------------------------------------------------------------------------
        public void SaveSignatoryRRData(string receivedBy, string checkedBy)
        {
            string accessConnectionString = GetAccessConnectionString();

            try
            {
                using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
                {
                    connection.Open();

                    string selectQuery = "SELECT COUNT(*) FROM SignatoryForRR";
                    int rowCount;

                    using (OleDbCommand selectCommand = new OleDbCommand(selectQuery, connection))
                    {
                        rowCount = (int)selectCommand.ExecuteScalar();
                    }

                    string signatoryQuery = null;

                    if (rowCount > 0)
                    {
                        signatoryQuery = "UPDATE SignatoryForRR SET ReceivedBy = ?, CheckedBy = ?";
                    }
                    else
                    {
                        signatoryQuery = "INSERT INTO SignatoryForRR (ReceivedBy, CheckedBy) VALUES (?, ?)";
                    }

                    using (OleDbCommand signatoryCommand = new OleDbCommand(signatoryQuery, connection))
                    {
                        signatoryCommand.Parameters.AddWithValue("@ReceivedBy", receivedBy);
                        signatoryCommand.Parameters.AddWithValue("@CheckedBy", checkedBy);

                        int rowsAffected = signatoryCommand.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            Console.WriteLine("Data inserted/updated successfully.");
                        }
                        else
                        {
                            Console.WriteLine("No rows were affected.");
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while updating SignatoryForRR table: {ex.Message}");
            }
        }
        public (string ReceivedBy, string CheckedBy) RetrieveSignatoryRRData()
        {
            string receivedBy = null;
            string checkedBy = null;

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
                {
                    connection.Open();

                    string query = "SELECT TOP 1 ReceivedBy, CheckedBy FROM SignatoryForRR";

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                receivedBy = reader["ReceivedBy"].ToString();
                                checkedBy = reader["CheckedBy"].ToString();
                            }
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while retrieving signatory for rr data: {ex.Message}");
            }

            return (receivedBy, checkedBy);
        }
        /*public (string ReceivedBy, string CheckedBy) RetrieveAllSignatoryForRRData()
        {
            string receivedBy = null;
            string checkedBy = null;

            string accessConnectionString = GetAccessConnectionString();

            try
            {
                using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
                {
                    connection.Open();
                    string query = "SELECT TOP 1 ReceivedBy, CheckedBy FROM SignatoryForRR";

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                receivedBy = reader["ReceivedBy"].ToString();
                                checkedBy = reader["CheckedBy"].ToString();
                            }
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while retrieving all signatory data: {ex.Message}");
            }

            return (receivedBy, checkedBy);
        }*/
        // ----------------------------------------------------------------------------------------------

        public int GetSeriesNumberFromDatabase(string columnName)
        {
            string accessConnectionString = GetAccessConnectionString();

            int currentSeries = 1; // Default to 1 if no value is found
            string query = $"SELECT {columnName} FROM Series"; // Replace 'SeriesTable' with your actual table name

            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        object result = command.ExecuteScalar();
                        if (result != null && int.TryParse(result.ToString(), out int series))
                        {
                            currentSeries = series;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error fetching series number: {ex.Message}");
                }
            }

            return currentSeries;
        }

        public void IncrementSeriesNumberInDatabase(string columnName)
        {
            string accessConnectionString = GetAccessConnectionString();

            string query = $"UPDATE Series SET {columnName} = {columnName} + 1";

            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error updating series number: {ex.Message}");
                }
            }
        }

        public void UpdateManualSeriesNumber(string columnName, int seriesNumber)
        {
            string query = $"UPDATE Series SET {columnName} = @SeriesNumber";

            using (OleDbConnection connection = new OleDbConnection(GetAccessConnectionString()))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@SeriesNumber", seriesNumber);
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error updating series number: {ex.Message}");
                }
            }
        }

        public class AmountToWordsConverter
        {
            private static string[] units = { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine" };
            private static string[] teens = { "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen" };
            private static string[] tens = { "", "Ten", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };
            private static string[] thousandsGroups = { "", " Thousand", " Million", " Billion" };

            public static string Convert(double amount)
            {
                if (amount == 0)
                    return "Zero Pesos Only";

                if (amount < 0)
                    return "Negative amount, cannot convert to words";

                int pesos = (int)Math.Floor(amount);
                int centavos = (int)Math.Round((amount - pesos) * 100);

                string pesoWords = ConvertToWords(pesos);
                string centavoWords = ConvertToWords(centavos);

                string result = "";
                if (centavos > 0)
                {
                    result = pesoWords + " Pesos";
                    //result = pesoWords + " and " + centavos + "/100 Pesos Only"; kanan terrys
                    result += " and " + centavoWords + " Centavos Only";
                }
                else
                {
                    result = pesoWords + " Pesos Only";
                }

                return result;
            }

            private static string ConvertToWords(int number)
            {
                if (number == 0)
                    return "Zero";

                if (number < 0)
                    return "Negative " + ConvertToWords(Math.Abs(number));

                string words = "";

                for (int i = 0; number > 0; i++)
                {
                    if (number % 1000 != 0)
                    {
                        words = ConvertHundreds(number % 1000) + thousandsGroups[i] + " " + words;
                    }
                    number /= 1000;
                }

                return words.Trim();
            }

            private static string ConvertHundreds(int number)
            {
                string words = "";

                if (number >= 100)
                {
                    words += units[number / 100] + " Hundred ";
                    number %= 100;
                }

                if (number >= 10 && number <= 19)
                {
                    words += teens[number - 10] + " ";
                    number = 0;
                }

                if (number >= 20)
                {
                    words += tens[number / 10] + " ";
                    number %= 10;
                }

                if (number >= 1 && number <= 9)
                {
                    words += units[number] + " ";
                }

                return words;
            }
        }


        // Helper to map Company Name -> Column Name (e.g. "North Luzon" -> "NL_CV")
        private string GetIVPColumnName(string formType, string companyName)
        {
            string prefix = "";
            switch (companyName)
            {
                case "North Luzon": prefix = "NL"; break;
                case "South Luzon": prefix = "SL"; break;
                case "Visayas": prefix = "VIS"; break;
                case "Mindanao": prefix = "MIN"; break;
                case "Metro Manila": prefix = "MM"; break;

                // Exact names for the specific companies
                case "Iberica Verheilen Pharmaceuticals Group.": return $"IVP_{formType}";
                case "Verheilen Iberica HealthCare Company Inc.": return $"VIHC_{formType}";
                case "My Health Shield NutriPharm Inc.": return $"MHS_{formType}";
                case "Greenfloor Innovations Corporation": return $"GIC_{formType}";

                case "Central Luzon": prefix = "CL"; break;
                default: return "";
            }
            return $"{prefix}_{formType}";
        }

        public int GetSeriesNumberFromDatabase(string formType, string companyName)
        {
            int seriesNumber = 1;
            string targetColumn = GetIVPColumnName(formType, companyName);

            if (string.IsNullOrEmpty(targetColumn)) return 1;

            // TARGETING TABLE: CVIVPIncrement
            string query = $"SELECT [{targetColumn}] FROM CVIVPIncrement WHERE ID = 1";

            using (OleDbConnection connection = new OleDbConnection(GetAccessConnectionString()))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        object result = command.ExecuteScalar();
                        if (result != null && result != DBNull.Value)
                        {
                            seriesNumber = Convert.ToInt32(result);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error retrieving series: {ex.Message}");
                }
            }
            return seriesNumber;
        }

        public void UpdateManualSeriesNumber(string formType, int seriesNumber, string companyName)
        {
            string targetColumn = GetIVPColumnName(formType, companyName);

            if (string.IsNullOrEmpty(targetColumn)) return;

            // TARGETING TABLE: CVIVPIncrement
            string query = $"UPDATE CVIVPIncrement SET [{targetColumn}] = @SeriesNumber WHERE ID = 1";

            using (OleDbConnection connection = new OleDbConnection(GetAccessConnectionString()))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@SeriesNumber", seriesNumber);
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error updating series: {ex.Message}");
                }
            }
        }
    }
}

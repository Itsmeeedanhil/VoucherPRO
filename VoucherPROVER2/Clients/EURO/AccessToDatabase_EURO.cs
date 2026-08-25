using System;
using System.Data.OleDb;
using System.IO;

namespace VoucherPROVER2.Clients.EURO
{
    public class AccessToDatabase_EURO
    {
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
                        switch (choice)
                        {
                            case 1:
                                signatoryQuery = "UPDATE Signatory SET PreparedByName = ?, PreparedByPosition = ?";
                                break;
                            case 2:
                                signatoryQuery = "UPDATE Signatory SET ReviewedByName = ?, ReviewedByPosition = ?";
                                break;
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
                                signatoryQuery = "INSERT INTO Signatory (ApprovedByName, ApprovedByPosition) VALUES (?, ?)";
                                break;
                            case 4:
                                signatoryQuery = "INSERT INTO Signatory (ReceivedByName, ReceivedByPosition) VALUES (?, ?)";
                                break;
                            default:
                                break;
                        }
                    }

                    if (signatoryQuery != null)
                    {
                        using (OleDbCommand signatoryCommand = new OleDbCommand(signatoryQuery, connection))
                        {
                            signatoryCommand.Parameters.AddWithValue("@Name", name);
                            signatoryCommand.Parameters.AddWithValue("@Position", position);

                            signatoryCommand.ExecuteNonQuery();
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while updating signatory table: {ex.Message}");
            }
        }

        public (string Name, string Position) GetSignatoryData(int choice)
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
                    switch (choice)
                    {
                        case 1:
                            query = "SELECT TOP 1 PreparedByName, PreparedByPosition FROM Signatory";
                            break;
                        case 2:
                            query = "SELECT TOP 1 ReviewedByName, ReviewedByPosition FROM Signatory";
                            break;
                        case 3:
                            query = "SELECT TOP 1 ApprovedByName, ApprovedByPosition FROM Signatory";
                            break;
                        case 4:
                            query = "SELECT TOP 1 ReceivedByName, ReceivedByPosition FROM Signatory";
                            break;
                        default:
                            break;
                    }

                    if (query != null)
                    {
                        using (OleDbCommand command = new OleDbCommand(query, connection))
                        {
                            using (OleDbDataReader reader = command.ExecuteReader())
                            {
                                if (reader.Read())
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

        public (
            string PreparedByName, string PreparedByPosition,
            string ReviewedByName, string ReviewedByPosition,
            string RecommendingApprovalName, string RecommendingApprovalPosition,
            string ApprovedByName, string ApprovedByPosition,
            string ReceivedByName, string ReceivedByPosition
            ) RetrieveAllSignatoryData()
        {
            string preparedByName = null;
            string preparedByPosition = null;
            string reviewedByName = null;
            string reviewedByPosition = null;
            string recommendingApprovalName = "";
            string recommendingApprovalPosition = "";
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

                    string query = "SELECT TOP 1 " +
                        "PreparedByName, PreparedByPosition, " +
                        "ReviewedByName, ReviewedByPosition, " +
                        "ApprovedByName, ApprovedByPosition, " +
                        "ReceivedByName, ReceivedByPosition " +
                        "FROM Signatory";

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                preparedByName = reader["PreparedByName"].ToString();
                                preparedByPosition = reader["PreparedByPosition"].ToString();

                                reviewedByName = reader["ReviewedByName"].ToString();
                                reviewedByPosition = reader["ReviewedByPosition"].ToString();

                                approvedByName = reader["ApprovedByName"].ToString();
                                approvedByPosition = reader["ApprovedByPosition"].ToString();

                                receivedByName = reader["ReceivedByName"].ToString();
                                receivedByPosition = reader["ReceivedByPosition"].ToString();
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

            return (
                preparedByName, preparedByPosition,
                reviewedByName, reviewedByPosition,
                recommendingApprovalName, recommendingApprovalPosition,
                approvedByName, approvedByPosition,
                receivedByName, receivedByPosition
                );
        }

        private string GetEUROColumnName(string formType)
        {
            return $"EURO_{formType}";
        }

        public int GetSeriesNumberFromDatabase(string formType, string companyName)
        {
            int seriesNumber = 1;
            string targetColumn = GetEUROColumnName(formType);

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
                    Console.WriteLine($"Error retrieving series for {targetColumn}: {ex.Message}");
                }
            }
            return seriesNumber;
        }

        public void UpdateManualSeriesNumber(string formType, int seriesNumber, string companyName)
        {
            string targetColumn = GetEUROColumnName(formType);

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
                    Console.WriteLine($"Error updating series for {targetColumn}: {ex.Message}");
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

                return words.Trim();
            }
        }
    }
}


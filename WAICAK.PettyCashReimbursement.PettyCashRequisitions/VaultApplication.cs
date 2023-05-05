using System;
using System.Data.SqlClient;
using System.Diagnostics;
using MFiles.VAF;
using MFiles.VAF.Common;
using MFiles.VAF.Configuration;
using MFiles.VAF.Configuration.JsonAdaptor;
using MFilesAPI;

namespace WAICAK.PettyCashReimbursement.PettyCashRequisitions
{
	/// <summary>
	/// The entry point for this Vault Application Framework application.
	/// </summary>
	/// <remarks>Examples and further information available on the developer portal: http://developer.m-files.com/. </remarks>
	public class VaultApplication
		: VaultApplicationBase
	{
        [StateAction("WFS.End of Requisition & Reconciliation")]
        public void CreatePettyCashReimbursement(StateEnvironment env)
        {

			SqlConnection cnn = new SqlConnection(@"Data Source=KENBOSRV01;Initial Catalog=PettyCashRequisitions;User ID=IntgrUsr;Password=passw0rd!");
			String query = null;
           // cnn = 
			string ptcashreqtitle = env.PropertyValues.SearchForProperty(1462).TypedValue.DisplayValue;
            int pcid = Int32.Parse(env.PropertyValues.SearchForProperty(1549).TypedValue.DisplayValue);
            string transactiondate = env.PropertyValues.SearchForProperty(1466).TypedValue.DisplayValue;
            //env.Objver.ID

            Lookups propExpenseLines = env.PropertyValues.SearchForProperty(1243).TypedValue.GetValueAsLookups();

            foreach( Lookup expenceitem in propExpenseLines)
            {
                string expenceiteminfo = expenceitem.DisplayValue;
                int spaceindex = expenceiteminfo.IndexOf(' ');
                int hyphenindex = expenceiteminfo.IndexOf('-');
                int lenght = hyphenindex - spaceindex - 1;
                string Expenselinecode = (expenceiteminfo.Substring(0, spaceindex));
                string expenselinedesc = (expenceiteminfo.Substring(spaceindex, lenght).Trim());
                Decimal propamt = Decimal.Parse(expenceiteminfo.Substring(hyphenindex + 1).Trim());
                //query = $@"INSERT INTO PCReimbursements (MFPREQID, TITLE, AMOUNTREQUISITIONED, TRANSACTIONDATE, EXCLUDEINSUMMARY, COMPILATIONCOMPLETE, EXPENSELINE, EXPENSELINEID) VALUES ({pcid}, '{ptcashreqtitle}' ,{propamt}, '{transactiondate}', 0, 0, '{expenselinedesc}', '{Expenselinecode}')";
                query = "INSERT INTO PCReimbursements (MFPREQID, TITLE, AMOUNTREQUISITIONED, TRANSACTIONDATE, EXCLUDEINSUMMARY, COMPILATIONCOMPLETE, EXPENSELINE, EXPENSELINEID)" +
                    "VALUES(@MFID, @NAMETITLE, @AMT, @TDATE, @EXCLUDESUMMARY, @CCOMPLETE, @EXPENSEDESC, @EXPLINEID)";
                SysUtils.ReportInfoToEventLog($"Query Statement: {query} ");
                SqlCommand command = new SqlCommand(query, cnn);
                command.Parameters.AddWithValue("@MFID", pcid);
                command.Parameters.AddWithValue("@NAMETITLE", ptcashreqtitle);
                command.Parameters.AddWithValue("@AMT", propamt);
                command.Parameters.AddWithValue("@TDATE", transactiondate);
                command.Parameters.AddWithValue("@EXCLUDESUMMARY", 0);
                command.Parameters.AddWithValue("@CCOMPLETE", 0);
                command.Parameters.AddWithValue("@EXPENSEDESC", expenselinedesc);
                command.Parameters.AddWithValue("@EXPLINEID", Expenselinecode);
                try
                {
                    cnn.Open();
                   // command.ExecuteNonQuery();
                    int aff = command.ExecuteNonQuery();
                    SysUtils.ReportInfoToEventLog($"Record Inserted Suceessfully: {aff} ");
                }
                catch (SqlException ex)
                {
                    SysUtils.ReportInfoToEventLog($"Error Generated. Details: {ex.ToString()} ");
                }
                finally
                {
                    cnn.Close();
                }

            }
           
            //Decimal propamt = Decimal.Parse(env.PropertyValues.SearchForProperty(1455).TypedValue.DisplayValue);			
			
            
        
			
            //string Expenselinecode = propExpenseLine.Substring(0, spaceindex);
            //SysUtils.ReportInfoToEventLog($"Expense line code: {Expenselinecode} ");
            //SysUtils.ReportInfoToEventLog($"Expense string length: {propExpenseLine.Length} ");
            //int newstringlength = propExpenseLine.Length - spaceindex;
            //string expenselinedesc = propExpenseLine.Substring(spaceindex, newstringlength).Trim();
            //SysUtils.ReportInfoToEventLog($"Expense line and amount: {expenselinedesc} ");
            
            
            
        }

    }
}
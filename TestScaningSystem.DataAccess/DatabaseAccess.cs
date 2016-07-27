using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;

namespace TestScaningSystem.DataAccess
{
    public class DatabaseAccess
    {
        SqlConnection conn;
        SqlCommand command;
        SqlDataAdapter adapter;
        SqlDataReader reader;
        DataTable dt;
        string query;
        public DatabaseAccess()
        {
            conn = new SqlConnection(@"Data Source=(local);Initial Catalog=QRTestingDB;Integrated Security=True");
        }

        #region GetStudentID
        public long GetStudentID(string[] info)
        {
            try
            {
                conn.Open();
                query = string.Format("SELECT S.StudentID FROM tblStudent S INNER JOIN tblStudentSubject SS ON S.StudentID = SS.StudentID INNER JOIN tblSubjects SJ ON SS.SubjectID = SJ.SubjectID WHERE S.StudentName = '{0}' AND S.StudentSurname = '{1}' AND SJ.SubjectCode = '{2}'", info[0], info[1], info[2]);
                command = new SqlCommand(query, conn);
                adapter = new SqlDataAdapter(command);
                dt = new DataTable();
                adapter.Fill(dt);
                long id = 0;
                foreach (DataRow row in dt.Rows)
                {
                    id = long.Parse(row["StudentId"].ToString());
                }
                return id;
            }
            catch (Exception)
            {
                return 0;
            }
            finally
            {
                conn.Close();
            }
        } 
        #endregion

        public bool Login(string username, string password)
        {
            try
            {
                conn.Open();
                query = string.Format("SELECT EmpID FROM tblEmployee WHERE EmpID = '{0}' AND EmpID ='{1}'", username, password);
                command = new SqlCommand(query, conn);
                adapter = new SqlDataAdapter(command);
                dt = new DataTable();
                adapter.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
                
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                conn.Close();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Collections;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using CrazyCoder.Class;


namespace CrazyCoder
{
    public partial class Form1 : Form
    {
        bool hasShowStructure = false;
        string databaseName = "";
        string tableName = "";
        string column = "";
        string code = "";
        string icode = "";
        string factory = "";
        string logic = "";
        string config = "";
        IList<ColumnInfo> columnInfos = new List<ColumnInfo>();

        public Form1()
        {
            InitializeComponent();
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
        }

        string GetConnectionString()
        {
            databaseName = lbDatabase.Items.Count > 0 ? lbDatabase.Items[lbDatabase.SelectedIndex].ToString() : "";
            string result = "Data Source=" + mtbServerIp.Text + ";Initial Catalog=" + databaseName + ";User ID=" + mtbDatabaseAccount.Text + ";Password=" + mtbDatabasePassword.Text;
            return result;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(GetConnectionString()))
                {
                    sqlConnection.Open();
                    SqlCommand sqlCommandSelectDatabase = new SqlCommand();
                    sqlCommandSelectDatabase.Connection = sqlConnection;
                    sqlCommandSelectDatabase.CommandText = "SELECT name FROM sys.databases";
                    SqlDataReader sqlDataReaderSelectDatabase = sqlCommandSelectDatabase.ExecuteReader();
                    if (sqlDataReaderSelectDatabase != null && sqlDataReaderSelectDatabase.HasRows)
                    {
                        lbDatabase.Items.Clear();
                        lbTable.Items.Clear();
                        while (sqlDataReaderSelectDatabase.Read())
                        {
                            lbDatabase.Items.Add(sqlDataReaderSelectDatabase["name"].ToString());
                        }
                    }
                    sqlDataReaderSelectDatabase.Dispose();
                    sqlCommandSelectDatabase.Dispose();
                    sqlConnection.Close();
                }
                lMessage.Text = "";
            }
            catch (Exception exception)
            {
                lbDatabase.Items.Clear();
                lbTable.Items.Clear();
                lMessage.Text = exception.Message;
            }
        }

        private void lbDatabase_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(GetConnectionString()))
                {
                    sqlConnection.Open();
                    SqlCommand sqlCommandSelectTable = new SqlCommand();
                    sqlCommandSelectTable.Connection = sqlConnection;
                    sqlCommandSelectTable.CommandText = "SELECT name FROM sysobjects WHERE type = 'U' ORDER BY name";
                    SqlDataReader sqlDataReaderSelectTable = sqlCommandSelectTable.ExecuteReader();
                    if (sqlDataReaderSelectTable != null && sqlDataReaderSelectTable.HasRows)
                    {
                        lbTable.Items.Clear();
                        while (sqlDataReaderSelectTable.Read())
                        {
                            lbTable.Items.Add(sqlDataReaderSelectTable["name"].ToString());
                        }
                    }
                    sqlDataReaderSelectTable.Dispose();
                    sqlCommandSelectTable.Dispose();
                    sqlConnection.Close();
                }
                lMessage.Text = "";
            }
            catch (Exception exception)
            {
                lMessage.Text = exception.Message;
            }
        }

        private void lbTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            hasShowStructure = false;
            ShowStructure();
            FillColumn();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string tagPageName = "tabPage" + (tabControl1.SelectedIndex + 1);
            string tagPageMethod = tagPageName + "_Foucs";

            Type form1Type = this.GetType();

            /*
            object[] parameters = new object[2];
            parameters[0] = sender;
            parameters[1] = e;
            
            Type[] types = new Type[parameters.Length];
            for (int i = 0; i < parameters.Length; i++)
            {
                types.SetValue(parameters[i].GetType(), i);
            }
            MethodInfo methodInfo = form1Type.GetMethod(tagPageMethod, types);
            //*/

            MethodInfo methodInfo = form1Type.GetMethod(tagPageMethod);

            if (methodInfo == null)
            {
                /*
                MessageBox.Show(form1Type.ToString());
                MessageBox.Show(tagPageMethod);
                MessageBox.Show("null");
                //*/
            }
            else
            {
                //methodInfo.Invoke(this, parameters);
                try
                {
                    methodInfo.Invoke(this, null);
                }
                catch
                {

                }
            }
        }

        public void tabPage1_Foucs()
        {
            //MessageBox.Show("tabPage1_Foucs");
        }

        public void tabPage2_Foucs()
        {
            //MessageBox.Show("tabPage2_Foucs");
            ShowStructure();
            hasShowStructure = true;
        }

        public void tabPage3_Foucs()
        {
            //MessageBox.Show("tabPage3_Foucs");
        }

        public void tabPage4_Foucs()
        {
            //MessageBox.Show("tabPage4_Foucs");
            FillCodeView();
        }

        public void tabPage5_Foucs()
        {
            //MessageBox.Show("tabPage5_Foucs");
        }

        private void FillColumn()
        {
            using (SqlConnection sqlConnection = new SqlConnection(GetConnectionString()))
            {
                tableName = lbTable.Items[lbTable.SelectedIndex].ToString();
                sqlConnection.Open();
                SqlCommand sqlCommandSelectColumn = new SqlCommand();
                sqlCommandSelectColumn.Connection = sqlConnection;
                sqlCommandSelectColumn.CommandText = @"SELECT syscolumns.name,systypes.name,syscolumns.length,syscomments.text,syscolumns.isnullable,sys.extended_properties.[value] FROM syscolumns join systypes on syscolumns.xusertype = systypes.xusertype left join syscomments on syscolumns.cdefault= syscomments.id 

LEFT OUTER JOIN sys.extended_properties ON syscolumns.id = sys.extended_properties.major_id AND syscolumns.colid = sys.extended_properties.minor_id AND sys.extended_properties.name = 'MS_Description' 

WHERE syscolumns.id = object_id('" + tableName + "')";
                SqlDataReader sqlDataReaderSelectColumn = sqlCommandSelectColumn.ExecuteReader();

                int i = 0;

                if (sqlDataReaderSelectColumn != null && sqlDataReaderSelectColumn.HasRows)
                {
                    tabPage3.Controls.Clear();

                    Label lColumnNameTmp = new Label();
                    lColumnNameTmp.Text = "列名";
                    lColumnNameTmp.Width = 200;
                    lColumnNameTmp.Height = 20;
                    lColumnNameTmp.Top = 20;
                    lColumnNameTmp.Left = 20;
                    tabPage3.Controls.Add(lColumnNameTmp);

                    Label lColumnTypeTmp = new Label();
                    lColumnTypeTmp.Text = "类型";
                    lColumnTypeTmp.Width = 80;
                    lColumnTypeTmp.Height = 20;
                    lColumnTypeTmp.Top = 20;
                    lColumnTypeTmp.Left =
lColumnNameTmp.Width
+
lColumnNameTmp.Left
;
                    tabPage3.Controls.Add(lColumnTypeTmp);

                    Label lColumnLengthTmp = new Label();
                    lColumnLengthTmp.Text = "长度";
                    lColumnLengthTmp.Width = 80;
                    lColumnLengthTmp.Height = 20;
                    lColumnLengthTmp.Top = 20;
                    lColumnLengthTmp.Left =
lColumnTypeTmp.Width
+
lColumnTypeTmp.Left
;
                    tabPage3.Controls.Add(lColumnLengthTmp);

                    Label lColumnDefaultValueTmp = new Label();
                    lColumnDefaultValueTmp.Text = "默认值";
                    lColumnDefaultValueTmp.Width = 100;
                    lColumnDefaultValueTmp.Height = 20;
                    lColumnDefaultValueTmp.Top = 20;
                    lColumnDefaultValueTmp.Left =
lColumnLengthTmp.Width
+
lColumnLengthTmp.Left
;
                    tabPage3.Controls.Add(lColumnDefaultValueTmp);

                    Label lColumnIsNullTmp = new Label();
                    lColumnIsNullTmp.Text = "是否空";
                    lColumnIsNullTmp.Width = 60;
                    lColumnIsNullTmp.Height = 20;
                    lColumnIsNullTmp.Top = 20;
                    lColumnIsNullTmp.Left =
lColumnDefaultValueTmp.Width
+
lColumnDefaultValueTmp.Left
;
                    tabPage3.Controls.Add(lColumnIsNullTmp);

                    columnInfos.Clear();

                    while (sqlDataReaderSelectColumn.Read())
                    {
                        ColumnInfo columnInfo = new ColumnInfo();
                        columnInfo.Name = sqlDataReaderSelectColumn[0].ToString();
                        columnInfo.Type = sqlDataReaderSelectColumn[1].ToString();
                        if (columnInfo.Type.ToLower() == "nvarchar")
                        {
                            //nvarchar，1个字符占2个字节varchar一个字符占1个字节，所以nvarchar用于中文，varchar用于数字英文特殊符号，范围1-4000，一个nvarchar字节长度为2，所以此处nvarchar(n)长度n=字节数/2，
                            int length = int.Parse(sqlDataReaderSelectColumn[2].ToString()) / 2;
                            columnInfo.Length = length.ToString();
                        }
                        else if (columnInfo.Type.ToLower() == "varchar")
                        {
                            //varchar，1个中文字符占用2个字节，一个英文，特殊符号字符占用1个字节，范围1-8000，所以此处varchar(n)长度n=字节数
                            int length = int.Parse(sqlDataReaderSelectColumn[2].ToString());
                            columnInfo.Length = sqlDataReaderSelectColumn[2].ToString();
                        }
                        else
                        {
                            columnInfo.Length = sqlDataReaderSelectColumn[2].ToString();
                        }
                        columnInfo.DefaultValue = sqlDataReaderSelectColumn[3].ToString();
                        columnInfo.IsNull = sqlDataReaderSelectColumn[4].ToString();
                        columnInfo.Description = sqlDataReaderSelectColumn[5].ToString();
                        columnInfo.Table = tableName;
                        columnInfos.Add(columnInfo);



                        Label lColumnName = new Label();
                        lColumnName.Text = columnInfo.Name;
                        lColumnName.Width = 200;
                        lColumnName.Height = 20;
                        lColumnName.Top = (i + 1) * 20 + 20;
                        lColumnName.Left = 20;
                        tabPage3.Controls.Add(lColumnName);

                        Label lColumnType = new Label();
                        lColumnType.Text = columnInfo.Type;
                        lColumnType.Width = 80;
                        lColumnType.Height = 20;
                        lColumnType.Top = (i + 1) * 20 + 20;
                        lColumnType.Left =
lColumnName.Width
+
lColumnName.Left
;
                        tabPage3.Controls.Add(lColumnType);

                        Label lColumnLength = new Label();
                        lColumnLength.Text = columnInfo.Length;
                        lColumnLength.Width = 80;
                        lColumnLength.Height = 20;
                        lColumnLength.Top = (i + 1) * 20 + 20;
                        lColumnLength.Left =
lColumnType.Width
+
lColumnType.Left
;
                        tabPage3.Controls.Add(lColumnLength);

                        Label lColumnDefaultValue = new Label();
                        lColumnDefaultValue.Text = columnInfo.DefaultValue;
                        lColumnDefaultValue.Width = 100;
                        lColumnDefaultValue.Height = 20;
                        lColumnDefaultValue.Top = (i + 1) * 20 + 20;
                        lColumnDefaultValue.Left =
lColumnLength.Width
+
lColumnLength.Left
;
                        tabPage3.Controls.Add(lColumnDefaultValue);

                        Label lColumnIsNull = new Label();
                        lColumnIsNull.Text = columnInfo.IsNull;
                        lColumnIsNull.Width = 60;
                        lColumnIsNull.Height = 20;
                        lColumnIsNull.Top = (i + 1) * 20 + 20;
                        lColumnIsNull.Left =
lColumnDefaultValue.Width
+
lColumnDefaultValue.Left
;
                        tabPage3.Controls.Add(lColumnIsNull);

                        i++;
                    }




                }
                sqlDataReaderSelectColumn.Dispose();
                sqlCommandSelectColumn.Dispose();
                sqlConnection.Close();

                tabControl1.SelectTab(1);
            }
        }

        private void FillCodeView()
        {
            if (rbSql.Checked)
            {
                CreateCodeBySql();
            }
            else
            {
                CreateCodeBySp();
            }
        }

        //生成sql版本DAL代码code
        #region
        private void CreateCodeBySql()
        {
            rtbCodeView.Text = "";
            code = "";

            code += @"using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Collections;
using System.Configuration;
using System.Data.SqlClient;
using " + mtbConfig.Text + @";
using " + mtbModel.Text + @";

namespace " + mtbDAL.Text + @"
{
    [Serializable]
    public partial class " + mtbDALName.Text + @" : IDisposable
    {
        //是否已经释放
        bool isDisposed = false;

        //手动调用或using超出作用域
        public void Dispose()
        {
            //手动调用或using超出作用域，调用true
            Dispose(true);
            //取消终结器，则不会调用析构函数，但是还是会由GC回收实例本身。
            GC.SuppressFinalize(this);	
        }

        void Dispose(bool isDisposedByManual)
        {
            lock(this)
            {
                if(!isDisposed)
                {
                    if(isDisposedByManual)
                    {
                        //释放托管资源
                    }
                    //释放非托管资源
                    //sqlConnection对象的释放资源比较特殊，一般不会直接释放资源，因为它是有个连接池的概念，连接池会自动调度空闲连接对象的，所以一般使用Close()即可。
                    //暂时关闭打开状态，sqlConnection对象还在的。
                    try
                    {
                        if (sqlConnection != null)
                        {
                            if (sqlConnection.State == ConnectionState.Open)
                            {
                                sqlConnection.Close();
                            }
                        }
                    }
                    catch
                    {

                    }
                    //暂时关闭打开状态，同时清空ConnectionString，sqlConnection对象还在的。
                    //sqlConnection.Dispose();
                    //推荐使用using关闭连接对象
                }
                isDisposed = true;
            }
        }

        //析构函数，当页面自动超出作用域，在垃圾回收之前，会调用析构函数
        ~" + mtbDALName.Text + @"()
        {
            Dispose(false);
        }

        SqlConnection sqlConnection = new SqlConnection(BaseConfig.ConnectionString);
        int _PageNumber = 1;
        int _PageSize = 10;
        int _RecordCount = 0;

        public int PageNumber
        {
            get{ return _PageNumber; }
            set{ _PageNumber = value; }
        }

        public int PageSize
        {
            get{ return _PageSize; }
            set{ _PageSize = value; }
        }

        public int PageCount
        {
            get{ return (_RecordCount - 1) / _PageSize + 1; }
        }

        public int RecordCount
        {
            get{ return _RecordCount; }
            set{ _RecordCount = value; }
        }

        //SqlCommand.Parameters与SqlParameter[]之间的转换
        SqlCommand CopySqlParametersToSqlCommandParameters(SqlCommand cmd, SqlParameter[] sqlParameters)
        {
            for(int i=0; i<sqlParameters.Length; i++)
            {
                cmd.Parameters.Add(new SqlParameter(sqlParameters[i].ToString(), sqlParameters[i].Value));
                cmd.Parameters[i].Direction = sqlParameters[i].Direction;
                cmd.Parameters[i].Size = sqlParameters[i].Size;
            }
            return cmd;
        }

        //获取IDataRecord对象并返回相应实体类对象
        " + mtbModelName.Text + @" GetInfo(IDataRecord dataRecord)
        {";
            code += @"
            " + mtbModelName.Text + @" obj = new " + mtbModelName.Text + @"();";
            for (int i = 0; i < columnInfos.Count; i++)
            {

                ColumnInfo columnInfo = columnInfos[i];
                if (columnInfo.Type == "varchar" || columnInfo.Type == "nvarchar" || columnInfo.Type == "text" || columnInfo.Type == "ntext")
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? """" : dataRecord[""" + columnInfo.Name + @"""].ToString();";
                }
                else if (columnInfo.Type == "datetime" || columnInfo.Type == "smalldatetime")
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? DateTime.Parse(""1900-01-01 00:00:00"") : Convert.ToDateTime(dataRecord[""" + columnInfo.Name + @"""].ToString());";
                }
                else if (columnInfo.Type == "bigint")
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? 0 : Convert.ToInt64(dataRecord[""" + columnInfo.Name + @"""].ToString());";
                }
                else if (columnInfo.Type == "int" || columnInfo.Type == "smallint" || columnInfo.Type == "tinyint")
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? 0 : Convert.ToInt32(dataRecord[""" + columnInfo.Name + @"""].ToString());";
                }
                else if (columnInfo.Type == "decimal")
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? 0 : Convert.ToDecimal(dataRecord[""" + columnInfo.Name + @"""].ToString());";
                }
                else if (columnInfo.Type == "bit")
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? false : Convert.ToBoolean(dataRecord[""" + columnInfo.Name + @"""].ToString());";
                }
                else
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? """" : dataRecord[""" + columnInfo.Name + @"""].ToString();";
                }
            }
            code += @"
            return obj;
        }
";

            if (cbSqlInsert.Checked)
            {
                code += @"
        /";
            }
            else
            {
                code += @"
        ";
            }

            code += @"/*/
        //新增单条记录方法
        //sql 语句
        public int Insert(" + mtbModelName.Text + @" obj)
        {
            sqlConnection.Open();
            SqlParameter[] sqlParameters = new SqlParameter[" + (columnInfos.Count - 1) + @"];";

            string spNameInsert = "SP_" + tableName + "_INSERT";

            for (int i = 1; i < columnInfos.Count; i++)
            {
                ColumnInfo columnInfo = columnInfos[i];
                code += @"
            sqlParameters[" + (i - 1) + @"] = new SqlParameter(""@" + columnInfo.Name + @""", obj." + columnInfo.Name + @");";
            }

            code += @"
            //string s = MSSQLHelper.ExecuteScalar(sqlConnection, System.Data.CommandType.StoredProcedure, " + spNameInsert + @", sqlParameters);
            string sql = """";
            SqlCommand cmdInsert = new SqlCommand();
            cmdInsert.Connection = sqlConnection;
            cmdInsert.CommandType = CommandType.Text;
            cmdInsert.CommandText = """";
            cmdInsert.CommandText += @""INSERT INTO [" + tableName + @"](";

            for (int i = 1; i < columnInfos.Count; i++)
            {
                ColumnInfo columnInfo = columnInfos[i];
                code += @"
                                        [" + columnInfo.Name + @"],";
                if (i == (columnInfos.Count - 1))
                {
                    code += @""";";
                }
            }

            code += @"
            cmdInsert.CommandText = cmdInsert.CommandText.Substring(0, cmdInsert.CommandText.Length - 1);
            cmdInsert.CommandText += @"") VALUES(";

            for (int i = 1; i < columnInfos.Count; i++)
            {
                ColumnInfo columnInfo = columnInfos[i];
                code += @"
                                        @" + columnInfo.Name + @",";
                if (i == (columnInfos.Count - 1))
                {
                    code += @""";";
                }
            }

            code += @"
            cmdInsert.CommandText = cmdInsert.CommandText.Substring(0, cmdInsert.CommandText.Length - 1);
            cmdInsert.CommandText += @"");SELECT @@IDENTITY;"";
            cmdInsert = CopySqlParametersToSqlCommandParameters(cmdInsert, sqlParameters);
            string s = cmdInsert.ExecuteScalar().ToString();
            cmdInsert.Parameters.Clear();
            cmdInsert.Dispose();
            int result = 0;
            if(s != """")
            {
                result = int.Parse(s);
            }
            sqlConnection.Close();
            return result;
        }
        //*/
";

            if (cbSqlUpdate.Checked)
            {
                code += @"
        /";
            }
            else
            {
                code += @"
        ";
            }

            code += @"/*/
        //修改单条记录方法
        //sql 语句
        public void Update(" + mtbModelName.Text + @" obj)
        {
            sqlConnection.Open();
            SqlParameter[] sqlParameters = new SqlParameter[" + columnInfos.Count + @"];";

            string spNameUpdate = "SP_" + tableName + "_UPDATE";

            for (int i = 0; i < columnInfos.Count; i++)
            {
                ColumnInfo columnInfo = columnInfos[i];
                code += @"
            sqlParameters[" + i + @"] = new SqlParameter(""@" + columnInfo.Name + @""", obj." + columnInfo.Name + @");";
            }

            code += @"
            //MSSQLHelper.ExecuteNonQuery(sqlConnection, System.Data.CommandType.StoredProcedure, " + spNameUpdate + @", sqlParameters);
            SqlCommand cmdUpdate = new SqlCommand();
            cmdUpdate.Connection = sqlConnection;
            cmdUpdate.CommandType = CommandType.Text;
            cmdUpdate.CommandText = """";
            cmdUpdate.CommandText += @""UPDATE [" + tableName + @"] SET ";

            for (int i = 1; i < columnInfos.Count; i++)
            {
                ColumnInfo columnInfo = columnInfos[i];
                code += @"
                                        [" + columnInfo.Name + @"]=@" + columnInfo.Name + @",";
                if (i == (columnInfos.Count - 1))
                {
                    code += @""";";
                }
            }
            code += @"
            cmdUpdate.CommandText = cmdUpdate.CommandText.Substring(0, cmdUpdate.CommandText.Length - 1);";
            ColumnInfo columnInfoUpdate = columnInfos[0];
            code += @"
            cmdUpdate.CommandText += "" WHERE " + columnInfoUpdate.Name + @"=@" + columnInfoUpdate.Name + @";"";
            cmdUpdate = CopySqlParametersToSqlCommandParameters(cmdUpdate, sqlParameters);
            cmdUpdate.ExecuteNonQuery();
            cmdUpdate.Parameters.Clear();
            cmdUpdate.Dispose();
            sqlConnection.Close();
        }
        //*/
";

            if (cbSqlDelete.Checked)
            {
                code += @"
        /";
            }
            else
            {
                code += @"
        ";
            }

            code += @"/*/
        //删除单条记录方法
        //sql 语句
        public void Delete(" + mtbModelName.Text + @" obj)
        {
            sqlConnection.Open();
            SqlParameter[] sqlParameters = new SqlParameter[1];";

            string spNameDelete = "SP_" + tableName + "_DELETE";


            ColumnInfo columnInfoDelete = columnInfos[0];
            code += @"
            sqlParameters[0] = new SqlParameter(""@" + columnInfoDelete.Name + @""", obj." + columnInfoDelete.Name + @");";

            code += @"
            //MSSQLHelper.ExecuteNonQuery(sqlConnection, System.Data.CommandType.StoredProcedure, " + spNameDelete + @", sqlParameters);
            SqlCommand cmdDelete = new SqlCommand();
            cmdDelete.Connection = sqlConnection;
            cmdDelete.CommandType = CommandType.Text;
            cmdDelete.CommandText = """";
            cmdDelete.CommandText += @""DELETE FROM [" + tableName + @"] WHERE " + columnInfoDelete.Name + @"=@" + columnInfoDelete.Name + @""";
            cmdDelete = CopySqlParametersToSqlCommandParameters(cmdDelete, sqlParameters);
            cmdDelete.ExecuteNonQuery();
            cmdDelete.Parameters.Clear();
            cmdDelete.Dispose();
            sqlConnection.Close();
        }
        //*/
";

            if (cbSqlSelect.Checked)
            {
                code += @"
        /";
            }
            else
            {
                code += @"
        ";
            }

            code += @"/*/
        //获取单条记录方法
        //sql 语句
        public " + mtbModelName.Text + @" Select(int id)
        {
            " + mtbModelName.Text + @" obj = new " + mtbModelName.Text + @"();
            sqlConnection.Open();
            SqlParameter[] sqlParameters = new SqlParameter[1];";

            string spNameSelect = "SP_" + tableName + "_SELECT";


            ColumnInfo columnInfoSelect = columnInfos[0];
            code += @"
            sqlParameters[0] = new SqlParameter(""@" + columnInfoSelect.Name + @""", id);";

            code += @"
            SqlCommand cmdSelect = new SqlCommand();
            cmdSelect.Connection = sqlConnection;
            cmdSelect.CommandType = CommandType.Text;
            cmdSelect.CommandText = """";
            cmdSelect.CommandText += @""SELECT * FROM [" + tableName + @"] WHERE " + columnInfoSelect.Name + @"=@" + columnInfoSelect.Name + @""";
            cmdSelect = CopySqlParametersToSqlCommandParameters(cmdSelect, sqlParameters);
            //SqlDataReader sdrSelect = MSSQLHelper.GetSqlDataReader(sqlConnection, System.Data.CommandType.StoredProcedure, " + spNameSelect + @", sps);
            SqlDataReader sdrSelect = cmdSelect.ExecuteReader();
            if(sdrSelect != null && sdrSelect.HasRows)
            {
                while(sdrSelect.Read())
                {";
            code += @"
                    obj = GetInfo(sdrSelect);
                }
            }
            else
            {
                obj = null;
            }
            sdrSelect.Dispose();
            cmdSelect.Parameters.Clear();
            cmdSelect.Dispose();
            sqlConnection.Close();
            return obj;
        }
        //*/
";

            if (cbSqlSelects.Checked)
            {
                code += @"
        /";
            }
            else
            {
                code += @"
        ";
            }

            code += @"/*/
        //获取多条记录方法
        //sql 语句
        public IList<" + mtbModelName.Text + @"> Selects()
        {
            IList<" + mtbModelName.Text + @"> objs = new List<" + mtbModelName.Text + @">();
            sqlConnection.Open();
            SqlParameter[] sqlParameters = new SqlParameter[3];";

            string spNameSelects = "SP_" + tableName + "_SELECTS";

            /*
            for (int i = 0; i < columnInfos.Count; i++)
            {
                ColumnInfo columnInfo = columnInfos[i];
                code += @"
    sqlParameters[" + i + @"] = new SqlParameter(""@" + columnInfo.Name + @""", obj." + columnInfo.Name + "@);";
            }
            //*/

            ColumnInfo columnInfoSelects = columnInfos[0];
            code += @"
            sqlParameters[0] = new SqlParameter(""@PageNumber"", PageNumber);
            sqlParameters[1] = new SqlParameter(""@PageSize"", PageSize);
            sqlParameters[2] = new SqlParameter(""@RecordCount"", RecordCount);
            sqlParameters[2].Direction = ParameterDirection.Output;
";

            code += @"
            SqlCommand cmdSelects = new SqlCommand();
            cmdSelects.Connection = sqlConnection;
            cmdSelects.CommandType = CommandType.Text;
            cmdSelects.CommandText = """";
            cmdSelects.CommandText += @""DECLARE @Min INT,@Max INT;
SET @Max = @PageSize * @PageNumber;
SET @Min = @PageSize * (@PageNumber - 1) + 1;
SELECT * FROM (



SELECT ROW_NUMBER() OVER(ORDER BY " + columnInfoSelects.Name + @" DESC) RANKING,* FROM [" + tableName + @"] WHERE 1 = 1



) AS T WHERE T.RANKING>=@Min AND T.RANKING<= @Max;



SELECT @RecordCount = COUNT(*) FROM " + tableName + @" WHERE 1 = 1"";
            //cmdSelects.CommandText += @""SELECT * FROM [" + tableName + @"] ORDER BY " + columnInfoSelects.Name + @" DESC"";
            cmdSelects = CopySqlParametersToSqlCommandParameters(cmdSelects, sqlParameters);
            //SqlDataReader sdrSelects = MSSQLHelper.GetSqlDataReader(sqlConnection, System.Data.CommandType.StoredProcedure, " + spNameSelects + @", sps);
            SqlDataReader sdrSelects = cmdSelects.ExecuteReader();
            if(sdrSelects != null && sdrSelects.HasRows)
            {
                while(sdrSelects.Read())
                {";
            code += @"
                    " + mtbModelName.Text + @" obj = GetInfo(sdrSelects);
                    objs.Add(obj);
                }
            }
            else
            {
                objs = null;
            }
            sdrSelects.Dispose();
            RecordCount = Convert.ToInt32(cmdSelects.Parameters[2].Value);
            cmdSelects.Parameters.Clear();
            cmdSelects.Dispose();
            sqlConnection.Close();
            return objs;
        }
        //*/
";

            code += @"
    }
}
";
            rtbCodeView.Text = code;
        }
        #endregion

        //生成存储过程版本DAL代码code
        #region
        private void CreateCodeBySp()
        {
            rtbCodeView.Text = "";
            code = "";

            code += @"using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Collections;
using System.Configuration;
using System.Data.SqlClient;
using " + mtbConfig.Text + @";
using " + mtbModel.Text + @";

namespace " + mtbDAL.Text + @"
{
    [Serializable]
    public partial class " + mtbDALName.Text + @"
    {
        SqlConnection sqlConnection = new SqlConnection(BaseConfig.ConnectionString);
        int _PageNumber = 1;
        int _PageSize = 10;
        int _RecordCount = 0;

        public int PageNumber
        {
            get{ return _PageNumber; }
            set{ _PageNumber = value; }
        }

        public int PageSize
        {
            get{ return _PageSize; }
            set{ _PageSize = value; }
        }

        public int PageCount
        {
            get{ return (_RecordCount - 1) / _PageSize + 1; }
        }

        public int RecordCount
        {
            get{ return _RecordCount; }
            set{ _RecordCount = value; }
        }

        //SqlCommand.Parameters与SqlParameter[]之间的转换
        public SqlCommand CopySqlParametersToSqlCommandParameters(SqlCommand cmd, SqlParameter[] sqlParameters)
        {
            for(int i=0; i<sqlParameters.Length; i++)
            {
                cmd.Parameters.Add(new SqlParameter(sqlParameters[i].ToString(), sqlParameters[i].Value));
                cmd.Parameters[i].Direction = sqlParameters[i].Direction;
                cmd.Parameters[i].Size = sqlParameters[i].Size;
            }
            return cmd;
        }

        //获取IDataRecord对象并返回相应实体类对象
        " + mtbModelName.Text + @" GetInfo(IDataRecord dataRecord)
        {";
            code += @"
            " + mtbModelName.Text + @" obj = new " + mtbModelName.Text + @"();";
            for (int i = 0; i < columnInfos.Count; i++)
            {

                ColumnInfo columnInfo = columnInfos[i];
                if (columnInfo.Type == "varchar" || columnInfo.Type == "nvarchar" || columnInfo.Type == "text" || columnInfo.Type == "ntext")
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? """" : dataRecord[""" + columnInfo.Name + @"""].ToString();";
                }
                else if (columnInfo.Type == "datetime" || columnInfo.Type == "smalldatetime")
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? DateTime.Parse(""1900-01-01 00:00:00"") : Convert.ToDateTime(dataRecord[""" + columnInfo.Name + @"""].ToString());";
                }
                else if (columnInfo.Type == "bigint")
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? 0 : Convert.ToInt64(dataRecord[""" + columnInfo.Name + @"""].ToString());";
                }
                else if (columnInfo.Type == "int" || columnInfo.Type == "smallint" || columnInfo.Type == "tinyint")
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? 0 : Convert.ToInt32(dataRecord[""" + columnInfo.Name + @"""].ToString());";
                }
                else if (columnInfo.Type == "decimal")
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? 0 : Convert.ToDecimal(dataRecord[""" + columnInfo.Name + @"""].ToString());";
                }
                else
                {
                    code += @"
            obj." + columnInfo.Name + @" = dataRecord[""" + columnInfo.Name + @"""] == DBNull.Value ? """" : dataRecord[""" + columnInfo.Name + @"""].ToString();";
                }
            }
            code += @"
            return obj;
        }
";


            if (cbSqlInsert.Checked)
            {
                code += @"
        /";
            }
            else
            {
                code += @"
        ";
            }

            code += @"/*/
        //新增单条记录方法
        //存储过程
        public int Insert(" + mtbModelName.Text + @" obj)
        {
            sqlConnection.Open();
            SqlParameter[] sqlParameters = new SqlParameter[" + (columnInfos.Count - 1) + @"];";

            string spNameInsert = "SP_" + tableName + "_INSERT";

            for (int i = 1; i < columnInfos.Count; i++)
            {
                ColumnInfo columnInfo = columnInfos[i];
                code += @"
            sqlParameters[" + (i - 1) + @"] = new SqlParameter(""@" + columnInfo.Name + @""", obj." + columnInfo.Name + @");";
            }

            code += @"
            //string s = MSSQLHelper.ExecuteScalar(sqlConnection, System.Data.CommandType.StoredProcedure, " + spNameInsert + @", sqlParameters);
            SqlCommand cmdInsert = new SqlCommand();
            cmdInsert.Connection = sqlConnection;
            cmdInsert.CommandType = CommandType.StoredProcedure;
            cmdInsert.CommandText = """ + spNameInsert + @""";
            cmdInsert = CopySqlParametersToSqlCommandParameters(cmdInsert, sqlParameters);
            string s = cmdInsert.ExecuteScalar().ToString();
            cmdInsert.Parameters.Clear();
            cmdInsert.Dispose();
            int result = 0;
            if(s != """")
            {
                result = int.Parse(s);
            }
            sqlConnection.Close();
            return result;
        }
        //*/
";



            if (cbSqlUpdate.Checked)
            {
                code += @"
        /";
            }
            else
            {
                code += @"
        ";
            }

            code += @"/*/
        //修改单条记录方法
        //存储过程
        public void Update(" + mtbModelName.Text + @" obj)
        {
            sqlConnection.Open();
            SqlParameter[] sqlParameters = new SqlParameter[" + columnInfos.Count + @"];";

            string spNameUpdate = "SP_" + tableName + "_UPDATE";

            for (int i = 0; i < columnInfos.Count; i++)
            {
                ColumnInfo columnInfo = columnInfos[i];
                code += @"
            sqlParameters[" + i + @"] = new SqlParameter(""@" + columnInfo.Name + @""", obj." + columnInfo.Name + @");";
            }

            code += @"
            //MSSQLHelper.ExecuteNonQuery(sqlConnection, System.Data.CommandType.StoredProcedure, " + spNameUpdate + @", sqlParameters);
            SqlCommand cmdUpdate = new SqlCommand();
            cmdUpdate.Connection = sqlConnection;
            cmdUpdate.CommandType = CommandType.StoredProcedure;
            cmdUpdate.CommandText = """ + spNameUpdate + @""";
            cmdUpdate = CopySqlParametersToSqlCommandParameters(cmdUpdate, sqlParameters);
            cmdUpdate.ExecuteNonQuery();
            cmdUpdate.Parameters.Clear();
            cmdUpdate.Dispose();
            sqlConnection.Close();
        }
        //*/
";



            if (cbSqlDelete.Checked)
            {
                code += @"
        /";
            }
            else
            {
                code += @"
        ";
            }

            code += @"/*/
        //删除单条记录方法
        //存储过程
        public void Delete(" + mtbModelName.Text + @" obj)
        {
            sqlConnection.Open();
            SqlParameter[] sqlParameters = new SqlParameter[1];";

            string spNameDelete = "SP_" + tableName + "_DELETE";


            ColumnInfo columnInfoDelete = columnInfos[0];
            code += @"
            sqlParameters[0] = new SqlParameter(""@" + columnInfoDelete.Name + @""", obj." + columnInfoDelete.Name + @");";

            code += @"
            //MSSQLHelper.ExecuteNonQuery(sqlConnection, System.Data.CommandType.StoredProcedure, " + spNameDelete + @", sqlParameters);
            SqlCommand cmdDelete = new SqlCommand();
            cmdDelete.Connection = sqlConnection;
            cmdDelete.CommandType = CommandType.StoredProcedure;
            cmdDelete.CommandText = """ + spNameDelete + @""";
            cmdDelete = CopySqlParametersToSqlCommandParameters(cmdDelete, sqlParameters);
            cmdDelete.ExecuteNonQuery();
            cmdDelete.Parameters.Clear();
            cmdDelete.Dispose();
            sqlConnection.Close();
        }
        //*/
";

            if (cbSqlSelect.Checked)
            {
                code += @"
        /";
            }
            else
            {
                code += @"
        ";
            }

            code += @"/*/
        //获取单条记录方法
        //存储过程
        public " + mtbModelName.Text + @" Select(int id)
        {
            " + mtbModelName.Text + @" obj = new " + mtbModelName.Text + @"();
            sqlConnection.Open();
            SqlParameter[] sqlParameters = new SqlParameter[1];";

            string spNameSelect = "SP_" + tableName + "_SELECT";


            ColumnInfo columnInfoSelect = columnInfos[0];
            code += @"
            sqlParameters[0] = new SqlParameter(""@" + columnInfoSelect.Name + @""", id);";

            code += @"
            SqlCommand cmdSelect = new SqlCommand();
            cmdSelect.Connection = sqlConnection;
            cmdSelect.CommandType = CommandType.StoredProcedure;
            cmdSelect.CommandText = """ + spNameSelect + @""";
            cmdSelect = CopySqlParametersToSqlCommandParameters(cmdSelect, sqlParameters);
            //SqlDataReader sdrSelect = MSSQLHelper.GetSqlDataReader(sqlConnection, System.Data.CommandType.StoredProcedure, " + spNameSelect + @", sps);
            SqlDataReader sdrSelect = cmdSelect.ExecuteReader();
            if(sdrSelect != null && sdrSelect.HasRows)
            {
                while(sdrSelect.Read())
                {";

            code += @"
                    obj = GetInfo(sdrSelect);
                }
            }
            else
            {
                obj = null;
            }
            sdrSelect.Dispose();
            cmdSelect.Parameters.Clear();
            cmdSelect.Dispose();
            sqlConnection.Close();
            return obj;
        }
        //*/
";



            if (cbSqlSelects.Checked)
            {
                code += @"
        /";
            }
            else
            {
                code += @"
        ";
            }

            code += @"/*/
        //获取多条记录方法
        //存储过程
        public IList<" + mtbModelName.Text + @"> Selects()
        {
            IList<" + mtbModelName.Text + @"> objs = new List<" + mtbModelName.Text + @">();
            sqlConnection.Open();
            SqlParameter[] sqlParameters = new SqlParameter[3];";

            string spNameSelects = "SP_" + tableName + "_SELECTS";

            /*
            for (int i = 0; i < columnInfos.Count; i++)
            {
                ColumnInfo columnInfo = columnInfos[i];
                code += @"
    sqlParameters[" + i + @"] = new SqlParameter(""@" + columnInfo.Name + @""", obj." + columnInfo.Name + "@);";
            }
            //*/

            code += @"
            sqlParameters[0] = new SqlParameter(""@PageNumber"", PageNumber);
            sqlParameters[1] = new SqlParameter(""@PageSize"", PageSize);
            sqlParameters[2] = new SqlParameter(""@RecordCount"", RecordCount);
            sqlParameters[2].Direction = ParameterDirection.Output;";

            code += @"
            SqlCommand cmdSelects = new SqlCommand();
            cmdSelects.Connection = sqlConnection;
            cmdSelects.CommandType = CommandType.StoredProcedure;
            cmdSelects.CommandText = """ + spNameSelects + @""";
            cmdSelects = CopySqlParametersToSqlCommandParameters(cmdSelects, sqlParameters);
            //SqlDataReader sdrSelects = MSSQLHelper.GetSqlDataReader(sqlConnection, System.Data.CommandType.StoredProcedure, " + spNameSelects + @", sps);
            SqlDataReader sdrSelects = cmdSelects.ExecuteReader();
            if(sdrSelects != null && sdrSelects.HasRows)
            {
                while(sdrSelects.Read())
                {";
            code += @"
                    " + mtbModelName.Text + @" obj = GetInfo(sdrSelects);
                    objs.Add(obj);
                }
            }
            else
            {
                objs = null;
            }
            sdrSelects.Dispose();
            RecordCount = Convert.ToInt32(cmdSelects.Parameters[2].Value);
            cmdSelects.Parameters.Clear();
            cmdSelects.Dispose();
            sqlConnection.Close();
            return objs;
        }
        //*/
";

            code += @"
    }
}
";

            rtbCodeView.Text = code;

        }
        #endregion

        //生成存储过程
        #region
        private void CreateSp()
        {
            using (SqlConnection sqlConnection = new SqlConnection(GetConnectionString()))
            {
                sqlConnection.Open();



                if (cbSqlInsert.Checked)
                {

                    //生成新增单条记录的存储过程
                    SqlCommand cmdInsert = new SqlCommand();
                    cmdInsert.Connection = sqlConnection;
                    string spName = "SP_" + tableName + "_INSERT";

                    //先删除已有存储过程
                    cmdInsert.CommandText = @"
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" + spName + @"]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[" + spName + @"]";
                    cmdInsert.ExecuteNonQuery();

                    cmdInsert.CommandText = @"
CREATE PROCEDURE " + spName + " ";

                    for (int i = 1; i < columnInfos.Count; i++)
                    {
                        ColumnInfo columnInfo = columnInfos[i];

                        cmdInsert.CommandText += "@" + columnInfo.Name + " " + columnInfo.Type;
                        if (columnInfo.Type == "nvarchar" || columnInfo.Type == "varchar")
                        {
                            cmdInsert.CommandText += "(" + columnInfo.Length + ")";
                        }

                        if (columnInfo.Type == "decimal")
                        {
                            cmdInsert.CommandText += "(18,2)";
                        }

                        cmdInsert.CommandText += ",";
                    }

                    cmdInsert.CommandText = cmdInsert.CommandText.Substring(0, cmdInsert.CommandText.Length - 1);
                    cmdInsert.CommandText += @"
AS
BEGIN
    INSERT INTO [" + tableName + "](";

                    for (int i = 1; i < columnInfos.Count; i++)
                    {
                        ColumnInfo columnInfo = columnInfos[i];
                        cmdInsert.CommandText += "[" + columnInfo.Name + "],";
                    }

                    cmdInsert.CommandText = cmdInsert.CommandText.Substring(0, cmdInsert.CommandText.Length - 1);
                    cmdInsert.CommandText += ") VALUES(";

                    for (int i = 1; i < columnInfos.Count; i++)
                    {
                        ColumnInfo columnInfo = columnInfos[i];
                        cmdInsert.CommandText += "@" + columnInfo.Name + ",";

                    }

                    cmdInsert.CommandText = cmdInsert.CommandText.Substring(0, cmdInsert.CommandText.Length - 1);
                    cmdInsert.CommandText += @")
    SELECT @@IDENTITY
END";
                    /*
                    MessageBox.Show(cmdInsert.CommandText);

                    StreamWriter sw = new StreamWriter("d:\\kaz.txt", false, Encoding.UTF8);
                    sw.Write(cmdInsert.CommandText);
                    sw.Dispose();
                    //*/

                    cmdInsert.ExecuteNonQuery();
                    cmdInsert.Dispose();
                }



                if (cbSqlUpdate.Checked)
                {
                    //生成修改单条记录的存储过程
                    SqlCommand cmdUpdate = new SqlCommand();
                    cmdUpdate.Connection = sqlConnection;
                    string spName = "SP_" + tableName + "_UPDATE";

                    //先删除已有存储过程
                    cmdUpdate.CommandText = @"
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" + spName + @"]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[" + spName + @"]";
                    cmdUpdate.ExecuteNonQuery();

                    cmdUpdate.CommandText = @"
CREATE PROCEDURE " + spName + @" ";

                    for (int i = 0; i < columnInfos.Count; i++)
                    {
                        ColumnInfo columnInfo = columnInfos[i];
                        cmdUpdate.CommandText += "@" + columnInfo.Name + " " + columnInfo.Type;
                        if (columnInfo.Type == "nvarchar" || columnInfo.Type == "varchar")
                        {
                            cmdUpdate.CommandText += "(" + columnInfo.Length + ")";
                        }

                        if (columnInfo.Type == "decimal")
                        {
                            cmdUpdate.CommandText += "(18,2)";
                        }

                        cmdUpdate.CommandText += ",";
                    }

                    cmdUpdate.CommandText = cmdUpdate.CommandText.Substring(0, cmdUpdate.CommandText.Length - 1);

                    cmdUpdate.CommandText += @"
AS
BEGIN
    UPDATE [" + tableName + @"] SET ";

                    for (int i = 1; i < columnInfos.Count; i++)
                    {
                        ColumnInfo columnInfo = columnInfos[i];
                        cmdUpdate.CommandText += "[" + columnInfo.Name + "]=@" + columnInfo.Name + ",";

                    }
                    cmdUpdate.CommandText = cmdUpdate.CommandText.Substring(0, cmdUpdate.CommandText.Length - 1);

                    ColumnInfo columnInfoUpdate = columnInfos[0];

                    cmdUpdate.CommandText += " WHERE " + columnInfoUpdate.Name + "=@" + columnInfoUpdate.Name;
                    cmdUpdate.CommandText += @"
END";
                    cmdUpdate.ExecuteNonQuery();
                    cmdUpdate.Dispose();
                }



                if (cbSqlDelete.Checked)
                {
                    //生成删除单条记录的存储过程
                    SqlCommand cmdDelete = new SqlCommand();
                    cmdDelete.Connection = sqlConnection;
                    string spName = "SP_" + tableName + "_DELETE";

                    //先删除已有存储过程
                    cmdDelete.CommandText = @"
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" + spName + @"]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[" + spName + @"]";
                    cmdDelete.ExecuteNonQuery();

                    cmdDelete.CommandText = @"
CREATE PROCEDURE " + spName + @" ";

                    ColumnInfo columnInfoDelete = columnInfos[0];

                    cmdDelete.CommandText += "@" + columnInfoDelete.Name + " " + columnInfoDelete.Type;
                    if (columnInfoDelete.Type == "nvarchar" || columnInfoDelete.Type == "varchar")
                    {
                        cmdDelete.CommandText += "(" + columnInfoDelete.Length + ")";
                    }

                    if (columnInfoDelete.Type == "decimal")
                    {
                        cmdDelete.CommandText += "(18,2)";
                    }

                    cmdDelete.CommandText += ",";


                    cmdDelete.CommandText = cmdDelete.CommandText.Substring(0, cmdDelete.CommandText.Length - 1);
                    cmdDelete.CommandText += @"
AS
BEGIN
    DELETE FROM [" + tableName + @"] WHERE " + columnInfoDelete.Name + @"=@" + columnInfoDelete.Name + @"
END";
                    cmdDelete.ExecuteNonQuery();
                    cmdDelete.Dispose();
                }



                if (cbSqlSelect.Checked)
                {
                    //生成获取单条记录的存储过程
                    SqlCommand cmdSelect = new SqlCommand();
                    cmdSelect.Connection = sqlConnection;
                    string spName = "SP_" + tableName + "_SELECT";

                    //先删除已有存储过程
                    cmdSelect.CommandText = @"
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" + spName + @"]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[" + spName + @"]";
                    cmdSelect.ExecuteNonQuery();

                    cmdSelect.CommandText = @"
CREATE PROCEDURE " + spName + @" ";

                    ColumnInfo columnInfoSelect = columnInfos[0];

                    cmdSelect.CommandText += "@" + columnInfoSelect.Name + " " + columnInfoSelect.Type;
                    if (columnInfoSelect.Type == "nvarchar" || columnInfoSelect.Type == "varchar")
                    {
                        cmdSelect.CommandText += "(" + columnInfoSelect.Length + ")";
                    }

                    if (columnInfoSelect.Type == "decimal")
                    {
                        cmdSelect.CommandText += "(18,2)";
                    }

                    cmdSelect.CommandText += ",";


                    cmdSelect.CommandText = cmdSelect.CommandText.Substring(0, cmdSelect.CommandText.Length - 1);
                    cmdSelect.CommandText += @"
AS
BEGIN
    SELECT * FROM [" + tableName + @"] WHERE " + columnInfoSelect.Name + @"=@" + columnInfoSelect.Name + @"
END";
                    cmdSelect.ExecuteNonQuery();
                    cmdSelect.Dispose();
                }




                if (cbSqlSelects.Checked)
                {
                    //生成获取多条记录的存储过程
                    SqlCommand cmdSelects = new SqlCommand();
                    cmdSelects.Connection = sqlConnection;
                    string spName = "SP_" + tableName + "_SELECTS";
                    cmdSelects.CommandText = @"
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" + spName + @"]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[" + spName + @"]";
                    cmdSelects.ExecuteNonQuery();

                    cmdSelects.CommandText = @"
CREATE PROCEDURE " + spName + @" @PageNumber INT,@PageSize INT,@RecordCount INT OUTPUT";

                    ColumnInfo columnInfoSelects = columnInfos[0];

                    /*
                    columnInfoSelects.CommandText += "@" + columnInfoSelects.Name + " " + columnInfoSelectSingleRecord.Type;
                    if (columnInfoSelects.Type == "nvarchar" || columnInfoSelects.Type == "varchar")
                    {
                        cmdSelects.CommandText += "(" + columnInfoSelects.Length + ")";
                    }

                    if (columnInfoSelects.Type == "decimal")
                    {
                        cmdSelects.CommandText += "(18,2)";
                    }

                    cmdSelects.CommandText += ",";
                    cmdSelects.CommandText = cmdSelects.CommandText.Substring(0, cmdSelects.CommandText.Length - 1);
                    //*/


                    cmdSelects.CommandText += @"
AS
BEGIN
    DECLARE @Min INT,@Max INT;
    SET @Max = @PageSize * @PageNumber;
    SET @Min = @PageSize * (@PageNumber - 1) + 1;
    SELECT * FROM (SELECT ROW_NUMBER() OVER(ORDER BY " + columnInfoSelects.Name + @" DESC) RANKING,* FROM [" + tableName + @"] WHERE 1 = 1) AS T WHERE T.RANKING>=@Min AND T.RANKING<= @Max;
    SELECT @RecordCount = COUNT(*) FROM " + tableName + @" WHERE 1 = 1;
END";
                    cmdSelects.ExecuteNonQuery();
                    cmdSelects.Dispose();

                }



                sqlConnection.Close();
            }
        }
        #endregion

        //生成实体类代码column
        #region
        private void ReadyForColumn()
        {

            //生成实体类
            column = "";
            column += @"using System;
using System.Collections.Generic;
using System.Text;

namespace " + mtbModel.Text + @"
{
    [Serializable]
    public class " + mtbModelName.Text + @"
    {";
            for (int j = 0; j < columnInfos.Count; j++)
            {
                ColumnInfo columnInfo = columnInfos[j];
                string columnType = "";

                switch (columnInfo.Type)
                {
                    case "varchar":
                        columnType = "string";
                        break;
                    case "nvarchar":
                        columnType = "string";
                        break;
                    case "text":
                        columnType = "string";
                        break;
                    case "ntext":
                        columnType = "string";
                        break;
                    case "bigint":
                        columnType = "long";
                        break;
                    case "int":
                        columnType = "int";
                        break;
                    case "smallint":
                        columnType = "int";
                        break;
                    case "tinyint":
                        columnType = "int";
                        break;
                    case "decimal":
                        columnType = "Decimal";
                        break;
                    case "datetime":
                        columnType = "DateTime";
                        break;
                    case "smalldatetime":
                        columnType = "DateTime";
                        break;
                    case "bit":
                        columnType = "bool";
                        break;
                    default:
                        columnType = "string";
                        break;
                }
                column += @"
        " + columnType + @" _" + columnInfo.Name + @";";
            }

            column += @"

        public " + mtbModelName.Text + @"()
        {
            //默认构造
            //生成时间 " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + @"
        }
";
            string columnType2 = "";
            string defaultValue = "";
            string nullValue = "";
            for (int j = 0; j < columnInfos.Count; j++)
            {
                ColumnInfo columnInfo = columnInfos[j];
                switch (columnInfo.Type)
                {
                    case "varchar":
                        columnType2 = "string";
                        defaultValue = "\"\"";
                        nullValue = "null";
                        break;
                    case "nvarchar":
                        columnType2 = "string";
                        defaultValue = "\"\"";
                        nullValue = "null";
                        break;
                    case "text":
                        columnType2 = "string";
                        defaultValue = "\"\"";
                        nullValue = "null";
                        break;
                    case "ntext":
                        columnType2 = "string";
                        defaultValue = "\"\"";
                        nullValue = "null";
                        break;
                    case "bigint":
                        columnType2 = "long";
                        defaultValue = "0";
                        nullValue = "0";
                        break;
                    case "int":
                        columnType2 = "int";
                        defaultValue = "0";
                        nullValue = "0";
                        break;
                    case "smallint":
                        columnType2 = "int";
                        defaultValue = "0";
                        nullValue = "0";
                        break;
                    case "tinyint":
                        columnType2 = "int";
                        defaultValue = "0";
                        nullValue = "0";
                        break;
                    case "decimal":
                        columnType2 = "Decimal";
                        defaultValue = "0";
                        nullValue = "0";
                        break;
                    case "datetime":
                        columnType2 = "DateTime";
                        defaultValue = "Convert.ToDateTime(\"1900-1-1\")";
                        nullValue = "Convert.ToDateTime(\"0001-1-1\")";
                        break;
                    case "smalldatetime":
                        columnType2 = "DateTime";
                        defaultValue = "Convert.ToDateTime(\"1900-1-1\")";
                        nullValue = "Convert.ToDateTime(\"0001-1-1\")";
                        break;
                    case "bit":
                        columnType2 = "bool";
                        defaultValue = "false";
                        nullValue = "false";
                        break;
                    default:
                        columnType2 = "string";
                        defaultValue = "\"\"";
                        nullValue = "null";
                        break;
                }
                column += @"
        /// <summary>
        /// " + columnInfo.Description + @"
        /// </summary>
        public " + columnType2 + @" " + columnInfo.Name + @"
        {
            get { return _" + columnInfo.Name + @" == " + nullValue + @" ? " + defaultValue + @" : _" + columnInfo.Name.Replace(" ", "_") + @"; }";
                if (columnInfo.Type == "varchar" || columnInfo.Type == "nvarchar")
                {
                    column += @"
            set { 
                    if(value.Length > " + columnInfo.Length + @")
                    {
                        _" + columnInfo.Name + @" = value.Substring(0, " + columnInfo.Length + @"); 
                    }
                    else
                    {
                        _" + columnInfo.Name + @" = value;
                    }
                }";
                }
                else
                {
                    column += @"
            set { _" + columnInfo.Name + @" = value; }";
                }
                column += @"
        }
";
            }

            column += @"    
    }
}
";
        }
        #endregion

        //生成业务逻辑代码logic
        #region
        protected void ReadyForBLL()
        {
            logic = "";
            logic += @"using System;
using System.Collections.Generic;
using System.Text;
using " + mtbIDAL.Text + @";
using " + mtbDALFactory.Text + @";
using " + mtbModel.Text + @";
using " + mtbDAL.Text + @";

namespace " + mtbBLL.Text + @"
{
    public class " + mtbDALName.Text + @"
    {
        I" + mtbDALName.Text + @" instance;
        object obj;
        string typeName = """ + mtbDALName.Text + @""";

        public " + mtbDALName.Text + @"()
        {
            obj = DataAccess.CreateInstance(typeName);
            instance = (I" + mtbDALName.Text + @")obj;
        }

        public int PageNumber
        {
            get { return instance.PageNumber; }
            set { instance.PageNumber = value; }
        }

        public int PageSize
        {
            get { return instance.PageSize; }
            set { instance.PageSize = value; }
        }

        public int RecordCount
        {
            get { return instance.RecordCount; }
            set { instance.RecordCount = value; }
        }

        public int Insert(" + mtbModelName.Text + @" obj)
        {
            return instance.Insert(obj);
        }

        public void Update(" + mtbModelName.Text + @" obj)
        {
            instance.Update(obj);
        }

        public void Delete(" + mtbModelName.Text + @" obj)
        {
            instance.Delete(obj);
        }

        public " + mtbModelName.Text + @" Select(int id)
        {
            return instance.Select(id);
        }

        public IList<" + mtbModelName.Text + @"> Selects()
        {
            return instance.Selects();
        }
    }
}
";
        }
        #endregion

        //生成数据访问工厂代码factory
        #region
        protected void ReadyForDALFactory()
        {
            factory = "";
            factory += @"using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Configuration;
using " + mtbConfig.Text + @";

namespace Test.DALFactory
{
    public class DataAccess
    {
        public static readonly string DAL = BaseConfig." + mtbDALString.Text + @";

        public static object CreateInstance(string className)
        {
            return Assembly.Load(DAL).CreateInstance(GetTypeName(className));
        }

        public static string GetTypeName(string className)
        {
            string fullClassName = DAL + ""."" + className;
            return fullClassName;
        }
    }
}
";
        }
        #endregion

        //生成数据访问接口代码icode
        #region
        protected void ReadyForIDAL()
        {
            icode = "";
            icode += @"using System;
using System.Collections.Generic;
using System.Text;
using " + mtbModel.Text + @";

namespace " + mtbIDAL.Text + @"
{
    public interface I" + mtbDALName.Text + @"
    {
        int PageNumber { get; set; }
        int PageSize { get; set; }
        int RecordCount { get; set; }
        int Insert(" + mtbModelName.Text + @" obj);
        void Update(" + mtbModelName.Text + @" obj);
        void Delete(" + mtbModelName.Text + @" obj);
        " + mtbModelName.Text + @" Select(int id);
        IList<" + mtbModelName.Text + @"> Selects();
    }
}
";
        }
        #endregion

        //生成配置代码
        #region
        protected void ReadyForConfig()
        {
            config = "";
            config += @"using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace " + mtbNameSpace.Text + @".Common
{
    public class BaseConfig
    {
        public static readonly string DAL = ConfigurationManager.AppSettings[""" + mtbDALString.Text + @"""] == null ? """" : ConfigurationManager.AppSettings[""" + mtbDALString.Text + @"""];
        public static readonly string ConnectionString = ConfigurationManager.AppSettings[""" + mtbConnectionString.Text + @"""] == null ? """" : ConfigurationManager.AppSettings[""" + mtbConnectionString.Text + @"""];
    }
}
";
        }
        #endregion

        private void rbSql_CheckedChanged(object sender, EventArgs e)
        {
            FillCodeView();
        }

        private void rbSp_CheckedChanged(object sender, EventArgs e)
        {
            FillCodeView();
        }

        private void cbSqlInsert_CheckedChanged(object sender, EventArgs e)
        {
            FillCodeView();
        }

        private void cbSqlUpdate_CheckedChanged(object sender, EventArgs e)
        {
            FillCodeView();
        }

        private void cbSqlDelete_CheckedChanged(object sender, EventArgs e)
        {
            FillCodeView();
        }

        private void cbSqlSelectSingleRecord_CheckedChanged(object sender, EventArgs e)
        {
            FillCodeView();
        }

        private void cbSqlSelectRecords_CheckedChanged(object sender, EventArgs e)
        {
            FillCodeView();
        }

        private void rbL2_CheckedChanged(object sender, EventArgs e)
        {
            ShowStructure();
        }

        private void rbL3_CheckedChanged(object sender, EventArgs e)
        {
            ShowStructure();
        }

        private void ShowStructure()
        {
            if (rbL2.Checked)
            {
                ShowL2();
            }
            else
            {
                ShowL3();
            }
        }

        private void ShowL2()
        {
            FillL();
            lBLL.Enabled = false;
            lBLLName.Enabled = false;
            lIDAL.Enabled = false;
            lDALFactory.Enabled = false;
            mtbBLL.Enabled = false;
            mtbBLLName.Enabled = false;
            mtbIDAL.Enabled = false;
            mtbDALFactory.Enabled = false;
        }

        private void ShowL3()
        {
            FillL();
            lBLL.Enabled = true;
            lBLLName.Enabled = true;
            lIDAL.Enabled = true;
            lDALFactory.Enabled = true;
            mtbBLL.Enabled = true;
            mtbBLLName.Enabled = true;
            mtbIDAL.Enabled = true;
            mtbDALFactory.Enabled = true;
        }

        private void FillL()
        {
            if (!hasShowStructure)
            {
                if (columnInfos.Count > 0)
                {
                    mtbNameSpace.Text = databaseName;
                    mtbModel.Text = mtbNameSpace.Text + ".Model";
                    mtbModelName.Text = columnInfos[0].Table + "Info";
                    mtbBLL.Text = mtbNameSpace.Text + ".BLL";
                    mtbBLLName.Text = columnInfos[0].Table;
                    mtbIDAL.Text = mtbNameSpace.Text + ".IDAL";
                    mtbDALFactory.Text = mtbNameSpace.Text + ".DALFactory";
                    mtbDAL.Text = mtbNameSpace.Text + ".DAL";
                    mtbDALName.Text = columnInfos[0].Table;
                    mtbConfig.Text = mtbNameSpace.Text + ".Common";
                    mtbDALString.Text = "DAL";
                    mtbConnectionString.Text = "ConnectionString";
                }
            }
        }

        private void mtbNameSpace_TextChanged(object sender, EventArgs e)
        {
            mtbModel.Text = mtbNameSpace.Text + ".Model";
            mtbModelName.Text = columnInfos[0].Table + "Info";
            mtbBLL.Text = mtbNameSpace.Text + ".BLL";
            mtbBLLName.Text = columnInfos[0].Table;
            mtbIDAL.Text = mtbNameSpace.Text + ".IDAL";
            mtbDALFactory.Text = mtbNameSpace.Text + ".DALFactory";
            mtbDAL.Text = mtbNameSpace.Text + ".DAL";
            mtbDALName.Text = columnInfos[0].Table;
            mtbConfig.Text = mtbNameSpace.Text + ".Common";
            FillCodeView();
        }

        private void lBrowse_Click(object sender, EventArgs e)
        {
            if (fbdFolder.ShowDialog() == DialogResult.OK)
            {
                mtbFolder.Text = fbdFolder.SelectedPath == "" ? fbdFolder.RootFolder.ToString() : fbdFolder.SelectedPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (rbSp.Checked)
                {
                    CreateSp();
                }

                ReadyForColumn();

                if (rbL2.Checked)
                {
                    ReadyForConfig();

                    //生成Model
                    Directory.CreateDirectory(mtbFolder.Text + "/" + mtbModel.Text);
                    StreamWriter streamWriterColumn = new StreamWriter(mtbFolder.Text + "/" + mtbModel.Text + "/" + mtbModelName.Text + ".cs", false, Encoding.UTF8);
                    streamWriterColumn.Write(column);
                    streamWriterColumn.Dispose();

                    //生成DAL
                    Directory.CreateDirectory(mtbFolder.Text + "/" + mtbDAL.Text);
                    StreamWriter streamWriterCode = new StreamWriter(mtbFolder.Text + "/" + mtbDAL.Text + "/" + mtbDALName.Text + ".cs", false, Encoding.UTF8);
                    streamWriterCode.Write(code);
                    streamWriterCode.Dispose();

                    //生成Config
                    Directory.CreateDirectory(mtbFolder.Text + "/" + mtbConfig.Text);
                    StreamWriter streamWriterConfig = new StreamWriter(mtbFolder.Text + "/" + mtbConfig.Text + "/BaseConfig.cs", false, Encoding.UTF8);
                    streamWriterConfig.Write(config);
                    streamWriterConfig.Dispose();
                }
                else
                {
                    ReadyForIDAL();
                    ReadyForDALFactory();
                    ReadyForBLL();
                    ReadyForConfig();

                    if (code.IndexOf(@"using " + mtbIDAL.Text + @";
using " + mtbModel.Text + @";") < 0)
                    {
                        code = code.Replace(@"using " + mtbModel.Text + @";", @"using " + mtbIDAL.Text + @";
using " + mtbModel.Text + @";");
                        code = code.Replace(@"public class " + mtbDALName.Text + @"", @"public class " + mtbDALName.Text + @" : I" + mtbDALName.Text + @"");
                    }
                    //生成Model
                    Directory.CreateDirectory(mtbFolder.Text + "/" + mtbModel.Text);
                    StreamWriter streamWriterColumn = new StreamWriter(mtbFolder.Text + "/" + mtbModel.Text + "/" + mtbModelName.Text + ".cs", false, Encoding.UTF8);
                    streamWriterColumn.Write(column);
                    streamWriterColumn.Dispose();

                    //生成BLL
                    Directory.CreateDirectory(mtbFolder.Text + "/" + mtbBLL.Text);
                    StreamWriter streamWriterLogic = new StreamWriter(mtbFolder.Text + "/" + mtbBLL.Text + "/" + mtbBLLName.Text + ".cs", false, Encoding.UTF8);
                    streamWriterLogic.Write(logic);
                    streamWriterLogic.Dispose();

                    //生成IDAL
                    Directory.CreateDirectory(mtbFolder.Text + "/" + mtbIDAL.Text);
                    StreamWriter streamWriterICode = new StreamWriter(mtbFolder.Text + "/" + mtbIDAL.Text + "/I" + mtbDALName.Text + ".cs", false, Encoding.UTF8);
                    streamWriterICode.Write(icode);
                    streamWriterICode.Dispose();

                    //生成DALFactory
                    Directory.CreateDirectory(mtbFolder.Text + "/" + mtbDALFactory.Text);
                    StreamWriter streamWriterFactory = new StreamWriter(mtbFolder.Text + "/" + mtbDALFactory.Text + "/DataAccess.cs", false, Encoding.UTF8);
                    streamWriterFactory.Write(factory);
                    streamWriterFactory.Dispose();

                    //生成DAL
                    Directory.CreateDirectory(mtbFolder.Text + "/" + mtbDAL.Text);
                    StreamWriter streamWriterCode = new StreamWriter(mtbFolder.Text + "/" + mtbDAL.Text + "/" + mtbDALName.Text + ".cs", false, Encoding.UTF8);
                    streamWriterCode.Write(code);
                    streamWriterCode.Dispose();

                    //生成Config
                    Directory.CreateDirectory(mtbFolder.Text + "/" + mtbConfig.Text);
                    StreamWriter streamWriterConfig = new StreamWriter(mtbFolder.Text + "/" + mtbConfig.Text + "/BaseConfig.cs", false, Encoding.UTF8);
                    streamWriterConfig.Write(config);
                    streamWriterConfig.Dispose();
                }

                //生成Xml
                if (cbGenerateXml.Checked)
                {
                    string xml = "";
                    xml += @"<?xml version=""1.0"" encoding=""utf-8""?>
<Root>
    <" + mtbDALName.Text + @">
        <" + mtbModelName.Text + @">";

                    for (int i = 0; i < columnInfos.Count; i++)
                    {
                        ColumnInfo columnInfo = columnInfos[i];
                        xml += @"
            <" + columnInfo.Name + @"><![CDATA[]]></" + columnInfo.Name + @">";
                    }

                    xml += @"
        </" + mtbModelName.Text + @">
    </" + mtbDALName.Text + @">
</Root>";

                    Directory.CreateDirectory(mtbFolder.Text + "/" + mtbNameSpace.Text + ".Xml");
                    StreamWriter streamWriterXml = new StreamWriter(mtbFolder.Text + "/" + mtbNameSpace.Text + ".Xml" + "/" + mtbDALName.Text + ".xml", false, Encoding.UTF8);
                    streamWriterXml.Write(xml);
                    streamWriterXml.Dispose();
                }

                lGenerate.Text = "成功生成";
            }
            catch (Exception ex)
            {
                lGenerate.Text = ex.Message;
            }
        }

        /// <summary>
        /// 默认显示code模式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenerateCodeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
        }

        /// <summary>
        /// 显示note模式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenerateNoteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = true;
            panel3.Visible = false;
            panel4.Visible = false;
        }

        /// <summary>
        /// 显示image模式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenerateImageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = true;
        }

        /// <summary>
        /// 显示about模式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AboutCrazyCoderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = true;
            panel4.Visible = false;
        }

        private void lNoteBrowser_Click(object sender, EventArgs e)
        {
            if (fbdNoteFolder.ShowDialog() == DialogResult.OK)
            {
                mtbNoteFolder.Text = fbdNoteFolder.SelectedPath == "" ? fbdNoteFolder.RootFolder.ToString() : fbdNoteFolder.SelectedPath;
            }
        }

        private void bGenerate_Click(object sender, EventArgs e)
        {
            //引用NPOI
            //引用NPOI.HPSF
            //引用NPOI.HSSF
            //引用NPOI.POIFS
            //引用NPOI.Util

            HSSFWorkbook workbook = new HSSFWorkbook();
            string nameSpace = "note";

            string[] files = Directory.GetFiles(mtbNoteFolder.Text);
            for (int i = 0; i < files.Length; i++)
            {
                StreamReader streamReader = File.OpenText(files[i]);
                string fileContent = streamReader.ReadToEnd();

                string[] fileParts = files[i].Split('\\');
                string fileName = fileParts[fileParts.Length - 1];
                if (fileName.IndexOf(".cs") >= 0)
                {
                    Match matchNameSpace = Regex.Match(fileContent, @"namespace ([^=\t\r\n{}]+)");
                    if (matchNameSpace != null)
                    {
                        nameSpace = matchNameSpace.Groups[1].Value;
                    }

                    //*
                    HSSFSheet sheet = workbook.CreateSheet(fileName);
                    sheet.SetColumnWidth(0, 256 * 40);
                    sheet.SetColumnWidth(1, 256 * 40);


                    //生成列名
                    HSSFRow rowHeader = sheet.CreateRow(0);
                    rowHeader.CreateCell(0).SetCellValue("方法名称");
                    rowHeader.CreateCell(1).SetCellValue("方法说明");
                    //*/

                    MatchCollection matchCollection = Regex.Matches(fileContent, @"(\/\/\/[^\{]*\>)[\s\t\r\n]*(public ([^\s]+\s)+[^\s]+\([^\)]*\))");
                    if (matchCollection != null)
                    {
                        HSSFPatriarch hssfPatriarch = sheet.CreateDrawingPatriarch();

                        for (int j = 0; j < matchCollection.Count; j++)
                        {
                            Match match = matchCollection[j];


                            //*
                            //生成数据

                            string functionName = match.Groups[2].Value;
                            functionName = "\r\n" + functionName;
                            string notationName = match.Groups[1].Value.Replace(" ", "");
                            //notationName = Regex.Replace(notationName, @"\>([^\r\n\<\>]+)", ">\r\n///$1");
                            //notationName = Regex.Replace(notationName, @"\<\/", "\r\n///</");
                            notationName = Regex.Replace(notationName, @"\>", ">\r\n///");
                            notationName = Regex.Replace(notationName, @"\/\/\/\r\n\/\/\/([^\r\n\<\>]+)", "///$1");
                            notationName = Regex.Replace(notationName, @"\<\/", "\r\n///</");
                            notationName = Regex.Replace(notationName, @"\/\/\/\r\n\/\/\/\<\/", "///</");
                            notationName = Regex.Replace(notationName, @"\>\r\n\/\/\/\<\/", ">\r\n///\r\n///</");
                            notationName = notationName.Replace("///", "");
                            notationName = "\r\n" + notationName;
                            //notationName = notationName.Substring(0, notationName.LastIndexOf("\r\n"));

                            int n = j + 1;
                            HSSFRow rowData = sheet.CreateRow(n);
                            //rowData.Height = 4000;
                            rowData.CreateCell(0).SetCellValue(functionName);
                            HSSFCellStyle hssfCellStyle = workbook.CreateCellStyle();
                            //hssfCellStyle.Rotation = 45;
                            hssfCellStyle.VerticalAlignment = 0;
                            hssfCellStyle.WrapText = true;
                            rowData.GetCell(0).CellStyle = hssfCellStyle;

                            //HSSFComment hssfComment0 = hssfPatriarch.CreateComment(new HSSFClientAnchor(0, 0, 0, 0, 1, 2, 4, 4));
                            //hssfComment0.String = new HSSFRichTextString(functionName);
                            //rowData.GetCell(0).CellComment = hssfComment0;


                            rowData.CreateCell(1).SetCellValue(notationName);
                            rowData.GetCell(1).CellStyle = hssfCellStyle;

                            //HSSFComment hssfComment1 = hssfPatriarch.CreateComment(new HSSFClientAnchor(0, 0, 0, 0, 1, 2, 4, 4));
                            //hssfComment1.String = new HSSFRichTextString(notationName);
                            //rowData.GetCell(1).CellComment = hssfComment1;
                            //*/

                        }
                    }
                }
            }






            string path = mtbNoteFolder.Text + @"\\" + nameSpace + @"." + DateTime.Now.ToString("yyyyMMddHHmmssfff") + @".xls";
            FileStream file = new FileStream(path, FileMode.Create);
            workbook.Write(file);
            file.Dispose();
            file.Close();

            System.Diagnostics.Process.Start(path);
        }

        private void lImageBrowser_Click(object sender, EventArgs e)
        {
            if (fbdImageFolder.ShowDialog() == DialogResult.OK)
            {
                mtbImageFolder.Text = fbdImageFolder.SelectedPath == "" ? fbdImageFolder.RootFolder.ToString() : fbdImageFolder.SelectedPath;
            }
        }

        private void bGenerateImage_Click(object sender, EventArgs e)
        {
            //mtbImageFolder.Text
            string[] filePaths = Directory.GetFiles(mtbImageFolder.Text);
            lImageTotal.Text = filePaths.Length.ToString();
            string newDirectory = "temp." + mudWidth.Value;
            if (!Directory.Exists(mtbImageFolder.Text + @"\" + newDirectory))
            {
                Directory.CreateDirectory(mtbImageFolder.Text + @"\" + newDirectory);
            }
            for (int i = 0; i < filePaths.Length; i++)
            {
                string originalFilePath = filePaths[i];
                string fileName = GetFileName(filePaths[i]);
                string newFilePath = filePaths[i].Replace(@"\" + fileName, @"\") + newDirectory + @"\" + fileName;
                ResizeImage(originalFilePath, newFilePath, Convert.ToInt16(mudWidth.Value));

                lImageDone.Text = (int.Parse(lImageDone.Text) + 1).ToString();
                System.Threading.Thread.Sleep(20);
                Application.DoEvents();
            }
        }

        protected string GetFileName(string filePath)
        {
            string fileName = "";
            string[] filePaths = filePath.Split('\\');
            if (filePaths.Length > 0)
            {
                fileName = filePaths[filePaths.Length - 1];
            }
            return fileName;
        }

        #region 定宽缩略程序
        protected string ResizeImage(string originalFilePath, string newFilePath, int width)
        {
            string result = "";
            System.Drawing.Image oImg = System.Drawing.Image.FromFile(originalFilePath);

            float HeightWidth = (float)oImg.Height / (float)oImg.Width;
            float WidthHeight = (float)oImg.Width / (float)oImg.Height;

            int img_width;
            int img_height;

            int x = 0;
            int y = 0;
            int new_width = width;
            int new_height = 0;

            if (oImg.Width > new_width)
            {
                img_width = new_width;
                img_height = (int)(img_width * HeightWidth);
            }
            else
            {
                img_width = oImg.Width;
                img_height = oImg.Height;
            }

            //if (img_height > new_height)
            //{
            //    img_height = new_height;
            //    img_width = (int)(new_height * WidthHeight);
            //}

            if (img_width < new_width)
            {
                x = (int)((new_width - img_width) / 2);
            }

            if (img_height < new_height)
            {
                y = (int)((new_height - img_height) / 2);
            }

            Bitmap bitmay = new Bitmap(new_width, img_height);
            Graphics g = Graphics.FromImage(bitmay);

            g.Clear(Color.FromName("white"));
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;


            g.DrawImage(oImg, new Rectangle(x, y, img_width, img_height), new Rectangle(0, 0, oImg.Width, oImg.Height), GraphicsUnit.Pixel);
            //if (File.Exists(Server.MapPath(mypath + filename + ".jpg")))
            //{
            //    File.Delete(Server.MapPath(mypath + filename + ".jpg"));
            //}

            //高质量jpg代码

            long[] quality = new long[1];
            quality[0] = 100;

            System.Drawing.Imaging.EncoderParameters encoderParams = new System.Drawing.Imaging.EncoderParameters();
            System.Drawing.Imaging.EncoderParameter encoderParam = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
            encoderParams.Param[0] = encoderParam;
            ImageCodecInfo[] arrayICI = ImageCodecInfo.GetImageEncoders();//获得包含有关内置图像编码解码器的信息的ImageCodecInfo 对象。
            ImageCodecInfo jpegICI = null;
            for (int i = 0; i < arrayICI.Length; i++)
            {
                if (arrayICI[i].FormatDescription.Equals("JPEG"))
                {
                    jpegICI = arrayICI[i];//设置JPEG编码
                    break;
                }
            }
            if (jpegICI != null)
            {
                //result = mypath + filename + "_" + width + "_" + height + ".jpg";
                bitmay.Save(newFilePath, jpegICI, encoderParams);
            }
            else
            {
                bitmay.Save(newFilePath, System.Drawing.Imaging.ImageFormat.Jpeg);
            }

            //高质量jpg代码 END

            //bitmay.Save(Server.MapPath(mypath + filename + "_mini.jpg"), System.Drawing.Imaging.ImageFormat.Jpeg);

            g.Dispose();
            bitmay.Dispose();

            oImg.Dispose();
            //result = result.Replace("~", "");
            return result;
        }
        #endregion

        #region 原始缩略程序
        protected string ResizeImage(string originalFilePath, string newFilePath, int width, int height)
        {
            string result = "";
            System.Drawing.Image oImg = System.Drawing.Image.FromFile(originalFilePath);

            float HeightWidth = (float)oImg.Height / (float)oImg.Width;
            float WidthHeight = (float)oImg.Width / (float)oImg.Height;

            int img_width;
            int img_height;

            int x = 0;
            int y = 0;
            int new_width = width;
            int new_height = height;

            if (oImg.Width > new_width)
            {
                img_width = new_width;
                img_height = (int)(new_width * HeightWidth);
            }
            else
            {
                img_width = oImg.Width;
                img_height = oImg.Height;
            }

            if (img_height > new_height)
            {
                img_height = new_height;
                img_width = (int)(new_height * WidthHeight);
            }

            if (img_width < new_width)
            {
                x = (int)((new_width - img_width) / 2);
            }

            if (img_height < new_height)
            {
                y = (int)((new_height - img_height) / 2);
            }

            Bitmap bitmay = new Bitmap(new_width, new_height);
            Graphics g = Graphics.FromImage(bitmay);

            g.Clear(Color.FromName("white"));
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;


            g.DrawImage(oImg, new Rectangle(x, y, img_width, img_height), new Rectangle(0, 0, oImg.Width, oImg.Height), GraphicsUnit.Pixel);
            //if (File.Exists(Server.MapPath(mypath + filename + ".jpg")))
            //{
            //    File.Delete(Server.MapPath(mypath + filename + ".jpg"));
            //}

            //高质量jpg代码

            long[] quality = new long[1];
            quality[0] = 100;

            System.Drawing.Imaging.EncoderParameters encoderParams = new System.Drawing.Imaging.EncoderParameters();
            System.Drawing.Imaging.EncoderParameter encoderParam = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
            encoderParams.Param[0] = encoderParam;
            ImageCodecInfo[] arrayICI = ImageCodecInfo.GetImageEncoders();//获得包含有关内置图像编码解码器的信息的ImageCodecInfo 对象。
            ImageCodecInfo jpegICI = null;
            for (int i = 0; i < arrayICI.Length; i++)
            {
                if (arrayICI[i].FormatDescription.Equals("JPEG"))
                {
                    jpegICI = arrayICI[i];//设置JPEG编码
                    break;
                }
            }
            if (jpegICI != null)
            {
                //result = mypath + filename + "_" + width + "_" + height + ".jpg";
                bitmay.Save(newFilePath, jpegICI, encoderParams);
            }
            else
            {
                bitmay.Save(newFilePath, System.Drawing.Imaging.ImageFormat.Jpeg);
            }

            //高质量jpg代码 END

            //bitmay.Save(Server.MapPath(mypath + filename + "_mini.jpg"), System.Drawing.Imaging.ImageFormat.Jpeg);

            g.Dispose();
            bitmay.Dispose();

            oImg.Dispose();
            //result = result.Replace("~", "");
            return result;
        }
        #endregion
    }
}

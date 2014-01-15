using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using NLog;

/// <summary>
/// FillTimeSheetSelect 的摘要描述
/// </summary>
public class GetTimeSheetSelect
{
    private static Logger saveLog = LogManager.GetLogger("GetTimeSheetSelect");
    public static SelectItems getOptions(string account)
    {
        SelectItems itemlist = new SelectItems();

        using (SqlCommand sqlCmd = new SqlCommand())
        {
            DataTable table = new DataTable();

            sqlCmd.CommandText = "SELECT ProjectID, ProjectName FROM Projects";
            table = SqlConnectTools.selectDataTable(sqlCmd);
            foreach (var item in table.AsEnumerable())
            {
                Option op = new Option();
                op.OptionId = item["ProjectID"].ToString();
                op.OptionName = item["ProjectName"].ToString();
                itemlist.Items.Proj.Add(op);
            }
            
            sqlCmd.CommandText = "SELECT t.TaskID StepID, t.TaskName StepName, ISNULL((SELECT TaskSort FROM TaskAccount WHERE Account=@Account AND Taskid = t.TaskID ), 0) TaskSort FROM Tasks t WHERE t.TaskTypeName = '專案階段'  ORDER BY TaskSort DESC";
            sqlCmd.Parameters.Add("Account", SqlDbType.VarChar, 20).Value = account;
            table = SqlConnectTools.selectDataTable(sqlCmd);
            foreach (var item in table.AsEnumerable())
            {
                Option op = new Option();
                op.OptionId = item["StepID"].ToString();
                op.OptionName = item["StepName"].ToString();
                op.OptionSort = Convert.ToInt32(item["TaskSort"]);
                itemlist.Items.Step.Add(op);
            }

            sqlCmd.CommandText = "SELECT t.TaskID , t.TaskName, ISNULL((SELECT TaskSort FROM TaskAccount WHERE Account=@Account AND TaskID = t.TaskID ), 0) TaskSort FROM Tasks t WHERE t.TaskTypeName = '工作類別'  ORDER BY TaskSort DESC";
            table = SqlConnectTools.selectDataTable(sqlCmd);
            foreach (var item in table.AsEnumerable())
            {
                Option op = new Option();
                op.OptionId = item["TaskID"].ToString();
                op.OptionName = item["TaskName"].ToString();
                op.OptionSort = Convert.ToInt32(item["TaskSort"]);
                itemlist.Items.Task.Add(op);
            }
        }

        return itemlist;
    }
}
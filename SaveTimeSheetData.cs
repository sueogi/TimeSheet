using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using Newtonsoft.Json;

/// <summary>
/// SaveTimeSheetData 的摘要描述
/// </summary>
public class SaveTimeSheetData
{
    public static string SaveData(SheetDetail sheet)
    {
        DateTime BeginTime = Convert.ToDateTime(sheet.SheetDate + " " + sheet.BeginHour + ":" + sheet.BeginMin);
        DateTime EndTime = Convert.ToDateTime(sheet.SheetDate + " " + sheet.EndHour + ":" + sheet.EndMin);

        string strCmd = string.Empty;
        if (string.IsNullOrEmpty(sheet.SheetID))
        {
            // 製作當日的 TimeSheet 開始跟結束時間字典檔
            string strCheckTime = "SELECT BeginTime, EndTime FROM Sheets WHERE (BeginTime BETWEEN @BeginTime AND @EndTime) AND Account = @Account";
            SqlCommand sqlCheckTime = new SqlCommand(strCheckTime);
            sqlCheckTime.Parameters.Add("BeginTime", SqlDbType.DateTime).Value = Convert.ToDateTime(sheet.SheetDate + " " + "00:00:00");
            sqlCheckTime.Parameters.Add("EndTime", SqlDbType.DateTime).Value = Convert.ToDateTime(sheet.SheetDate + " " + "23:59:59");
            sqlCheckTime.Parameters.Add("Account", SqlDbType.VarChar, 20).Value = sheet.Account;
            DataTable table = SqlConnectTools.selectDataTable(sqlCheckTime);
            Dictionary<DateTime, DateTime> timeDics = new Dictionary<DateTime, DateTime>();
            foreach (var item in table.AsEnumerable())
            {
                timeDics.Add(Convert.ToDateTime(item["BeginTime"]), Convert.ToDateTime(item["EndTime"]));
            }
            // 只要開始時間一樣就 failure
            if (timeDics.Keys.Contains(BeginTime))
            {
                return "false";
            }
            foreach (var item in timeDics)
            {
                // failure state1 : 字典檔 12:00~13:00 寫入 12:10~13:00
                if ((item.Key < BeginTime && item.Value > BeginTime) ||
                    // failure state2 : 字典檔 12:00~13:00 寫入 11:00~12:01
                    (item.Key > BeginTime && item.Key < EndTime))
                {
                    return "false";
                }
            }

            strCmd = "INSERT INTO Sheets (Account, BeginTime, EndTime, TaskID, Summary";
            strCmd += !sheet.ProjectID.Equals("0") ?
                ", ProjectID, StageID) VALUES (@Account, @BeginTime, @EndTime, @TaskID, @Summary, @ProjectID, @StageID)" :
                ") VALUES (@Account, @BeginTime, @EndTime, @TaskID, @Summary)";
        }
        else
        {
            strCmd = "UPDATE Sheets SET BeginTime = @BeginTime, EndTime = @EndTime, TaskID = @TaskID, Summary = @Summary";
            strCmd += !sheet.ProjectID.Equals("0") ? ", ProjectID = @ProjectID, StageID = @StageID" : string.Empty;
            strCmd += ", ModifyTime = @ModifyTime WHERE SheetID = @SheetID";
        }

        SqlCommand sqlCmd = new SqlCommand(strCmd);
        sqlCmd.Parameters.Add("Account", SqlDbType.VarChar, 20).Value = sheet.Account;
        sqlCmd.Parameters.Add("BeginTime", SqlDbType.DateTime).Value = BeginTime;
        sqlCmd.Parameters.Add("EndTime", SqlDbType.DateTime).Value = EndTime;
        sqlCmd.Parameters.Add("TaskID", SqlDbType.Int).Value = sheet.TaskID;
        sqlCmd.Parameters.Add("ProjectID", SqlDbType.Char, 7).Value = sheet.ProjectID;
        sqlCmd.Parameters.Add("StageID", SqlDbType.Int).Value = sheet.StepID;
        sqlCmd.Parameters.Add("Summary", SqlDbType.NVarChar, 50).Value = sheet.Summary;
        sqlCmd.Parameters.Add("ModifyTime", SqlDbType.DateTime).Value = DateTime.Now;
        Guid sheetId;

        Guid.TryParse(sheet.SheetID, out sheetId);

        sqlCmd.Parameters.Add("SheetID", SqlDbType.UniqueIdentifier).Value = sheetId;
        bool result = SqlConnectTools.setSqlCmd(sqlCmd);

        if (result && sheetId == (new Guid()))
        {
            string strFindId = "SELECT SheetID FROM Sheets WHERE Account = @Account AND BeginTime = @BeginTime";
            SqlCommand sqlFindId = new SqlCommand(strFindId);
            sqlFindId.Parameters.Add("Account", SqlDbType.VarChar, 20).Value = sheet.Account;
            sqlFindId.Parameters.Add("BeginTime", SqlDbType.DateTime).Value = BeginTime;
            sheetId = SqlConnectTools.getSheetGuid(sqlFindId);
        }


        string strAddSort = string.Empty;
        if (SqlConnectTools.checkTaskAccount(sheet.Account, sheet.TaskID))
        {
            strAddSort = "UPDATE TaskAccount SET TaskSort = (TaskSort + 1) WHERE TaskID = @TaskID AND Account = @Account";
        }
        else
        {
            strAddSort = "INSERT INTO TaskAccount (TaskID, Account, TaskSort) VALUES (@TaskID, @Account, 1)";
        }
        SqlCommand sqlAddTaskSort = new SqlCommand(strAddSort);
        sqlAddTaskSort.Parameters.Add("TaskID", SqlDbType.Int).Value = sheet.TaskID;
        sqlAddTaskSort.Parameters.Add("Account", SqlDbType.VarChar, 20).Value = sheet.Account;
        SqlConnectTools.setSqlCmd(sqlAddTaskSort);
        sqlAddTaskSort.Dispose();


        if (SqlConnectTools.checkTaskAccount(sheet.Account, sheet.StepID))
        {
            strAddSort = "UPDATE TaskAccount SET TaskSort = (TaskSort + 1) WHERE TaskID = @TaskID AND Account = @Account";
        }
        else
        {
            strAddSort = "INSERT INTO TaskAccount (TaskID, Account, TaskSort) VALUES (@TaskID, @Account, 1)";
        }
        SqlCommand sqlAddStepSort = new SqlCommand(strAddSort);
        sqlAddStepSort.Parameters.Add("TaskID", SqlDbType.Int).Value = sheet.StepID;
        sqlAddStepSort.Parameters.Add("Account", SqlDbType.VarChar, 20).Value = sheet.Account;
        SqlConnectTools.setSqlCmd(sqlAddStepSort);
        sqlAddStepSort.Dispose();



        // 寫入 Members 資料表
        if (!SqlConnectTools.checkMemberProject(sheet.Account, sheet.ProjectID))
        {
            string strCheckMember = "INSERT INTO Members (RoleID, Account, ProjectID) VALUES ('2', @Account, @ProjectID)";
            SqlCommand sqlCheckMember = new SqlCommand(strCheckMember);
            sqlCheckMember.Parameters.Add("Account", SqlDbType.VarChar, 20).Value = sheet.Account;
            sqlCheckMember.Parameters.Add("ProjectID", SqlDbType.Char, 7).Value = sheet.ProjectID;
            SqlConnectTools.setSqlCmd(sqlCheckMember);
            sqlCheckMember.Dispose();
        }

        return result ? sheetId.ToString() : "false";
    }
}
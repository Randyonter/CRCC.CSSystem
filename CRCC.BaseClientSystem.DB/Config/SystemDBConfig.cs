/****************************************************************************************************************************************************************
 *
 * 版权所有: 2023  All Rights Reserved.
 * 文 件 名:  SystemDBConfig
 * 描    述:  
 * 创 建 者：  Ltf
 * 创建时间：  2023/3/30 9:35:18
 * 历史更新记录:
 * =============================================================================================
*修改标记
*修改时间：2023/3/30 9:35:18
*修改人： Ltf
*版本号： V1.0.0.0
*描述：
 * =============================================================================================

****************************************************************************************************************************************************************/
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRCC.BaseClientSystem.DB.Config
{
    public static class SystemDBConfig
    {
        public static string DBConnectStr = "server=localhost;Database=platformdatabase;Uid=root;Pwd=root123;Allow User Variables=True;AllowLoadLocalInfile=true";

        public static SqlSugarScope Db = new SqlSugarScope(new ConnectionConfig()
        {
            DbType = DbType.MySql,//数据库类型
            ConnectionString = DBConnectStr,
            IsAutoCloseConnection = true
        },
        db =>
        {
            db.Aop.OnLogExecuting = (sql, p) =>
            {
                Console.WriteLine(sql);
                Console.WriteLine(string.Join(",", p?.Select(t => t.ParameterName + ":" + t.Value) ?? new List<string> { string.Empty }));
            };
        });
    }
}

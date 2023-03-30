using CRCC.BaseClientSystem.DB.DBSeed;
using CRCC.BaseClientSystem.IRepositories;
using CRCC.Model;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace CRCC.BaseClientSystem.Repositories
{
    public class Repository<TEntity> : BaseIRespositiories<TEntity> where TEntity : class, new()
    {
        private MyContext context;
        private SqlSugarScope db;
        private SimpleClient<TEntity> entityDB;

        public MyContext Context
        {
            get { return context; }
            set { context = value; }
        }

        internal SqlSugarScope Db
        {
            get { return db; }
            private set { db = value; }
        }

        internal SimpleClient<TEntity> EntityDB
        {
            get { return entityDB; }
            private set { entityDB = value; }
        }

        ISqlSugarClient BaseIRespositiories<TEntity>.Db => db;

        public Repository()
        {
            context = MyContext.GetDbContext();
            db = context.Db;
            entityDB = context.GetEntityDB<TEntity>(db);
        }

        /// <summary>
        /// 写入实体数据
        /// </summary>
        /// <param name="entity">实体对象</param>
        /// <returns>影响行数</returns>
        public async Task<int> Add(TEntity model)
        {
            var i = await Task.Run(() => db.Insertable(model).ExecuteCommandAsync());
            return i;
        }

        /// <summary>
        /// 批量插入实体(速度快)
        /// </summary>
        /// <param name="listEntity">实体集合</param>
        /// <returns>影响行数</returns>
        public async Task<int> Add(List<TEntity> listEntity)
        {
            return await db.Insertable(listEntity.ToArray()).ExecuteCommandAsync();
        }

        /// <summary>
        /// 批量插入实体(速度快)
        /// </summary>
        /// <param name="listEntity">实体集合</param>
        /// <returns>影响行数</returns>
        public async Task<int> AddBulkCopy(List<TEntity> listEntity)
        {
            return await db.Fastest<TEntity>().BulkCopyAsync(listEntity);
        }

        /// <summary>
        /// 批量插入实体(速度快)
        /// </summary>
        /// <param name="listEntity">实体集合</param>
        /// <returns>影响行数</returns>
        public async Task<int> UpdateList(List<TEntity> listEntity)
        {
            return await db.Fastest<TEntity>().BulkUpdateAsync(listEntity);
        }

        /// <summary>
        /// 根据实体删除一条数据
        /// </summary>
        /// <param name="entity">实体类</param>
        /// <returns></returns>
        public async Task<bool> Delete(TEntity model)
        {
            return await db.Deleteable(model).ExecuteCommandHasChangeAsync();
        }

        /// <summary>
        /// 删除指定ID的数据
        /// </summary>
        /// <param name="id">主键ID</param>
        /// <returns></returns>
        public async Task<bool> DeleteById(object id)
        {
            return await db.Deleteable<TEntity>().In(id).ExecuteCommandHasChangeAsync();
        }

        /// <summary>
        /// 删除指定ID集合的数据(批量删除)
        /// </summary>
        /// <param name="ids">主键ID集合</param>
        public async Task<bool> DeleteByIds(object[] ids)
        {
            return await db.Deleteable<TEntity>().In(ids).ExecuteCommandHasChangeAsync();
        }

        /// <summary>
        /// 删除指定ID的数据(用于存在多个主键）
        /// </summary>
        /// <param name="ids">主键ID集合</param>
        public async Task<bool> DeleteByWhere(Expression<Func<TEntity, bool>> expression)
        {
            return await db.Deleteable<TEntity>().Where(expression).ExecuteCommandHasChangeAsync();
        }

        /// <summary>
        /// 功能描述:查询所有数据
        /// </summary>
        /// <returns>数据列表</returns>
        public async Task<List<TEntity>> Query()
        {
            return await db.Queryable<TEntity>().ToListAsync();
        }

        /// <summary>
        /// 功能描述:查询数据列表
        /// </summary>
        /// <param name="where">条件</param>
        /// <returns>数据列表</returns>
        public async Task<List<TEntity>> Query(string where)
        {
            return await db.Queryable<TEntity>().WhereIF(!string.IsNullOrEmpty(where), where).ToListAsync();
        }

        /// <summary>
        /// 功能描述:查询数据列表
        /// </summary>
        /// <param name="whereExpression">whereExpression</param>
        /// <returns>数据列表</returns>
        public async Task<List<TEntity>> Query(Expression<Func<TEntity, bool>> whereExpression)
        {
            return await db.Queryable<TEntity>().WhereIF(whereExpression != null, whereExpression).ToListAsync();
        }

        /// <summary>
        /// 功能描述:按照特定列查询数据列表
        /// </summary>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="expression"></param>
        /// <returns></returns>
        public async Task<List<TResult>> Query<TResult>(Expression<Func<TEntity, TResult>> expression)
        {
            return await db.Queryable<TEntity>().Select(expression).ToListAsync();
        }

        /// <summary>
        /// 功能描述:按照特定列查询数据列表带条件排序
        /// </summary>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="whereExpression">过滤条件</param>
        /// <param name="expression">查询实体条件</param>
        /// <param name="orderByFields">排序条件</param>
        /// <returns></returns>
        public async Task<List<TResult>> Query<TResult>(Expression<Func<TEntity, TResult>> expression, Expression<Func<TEntity, bool>> whereExpression, string orderByFields)
        {
            return await db.Queryable<TEntity>().OrderByIF(!string.IsNullOrEmpty(orderByFields), orderByFields).WhereIF(whereExpression != null, whereExpression).Select(expression).ToListAsync();
        }

        /// <summary>
        /// 功能描述:查询一个列表
        /// </summary>
        /// <param name="whereExpression">条件表达式</param>
        /// <param name="orderByFields">排序字段，如name asc,age desc</param>
        /// <returns>数据列表</returns>
        public async Task<List<TEntity>> Query(Expression<Func<TEntity, bool>> whereExpression, string orderByFields)
        {
            return await db.Queryable<TEntity>().WhereIF(whereExpression != null, whereExpression).OrderByIF(orderByFields != null, orderByFields).ToListAsync();
        }

        /// <summary>
        /// 功能描述:查询一个列表
        /// </summary>
        /// <param name="whereExpression"></param>
        /// <param name="orderByExpression"></param>
        /// <param name="isAsc"></param>
        /// <returns></returns>
        public async Task<List<TEntity>> Query(Expression<Func<TEntity, bool>> whereExpression, Expression<Func<TEntity, object>> orderByExpression, bool isAsc = true)
        {
            return await db.Queryable<TEntity>().OrderByIF(orderByExpression != null, orderByExpression, isAsc ? OrderByType.Asc : OrderByType.Desc).WhereIF(whereExpression != null, whereExpression).ToListAsync();
        }

        /// <summary>
        /// 功能描述:查询一个列表
        /// </summary>
        /// <param name="where">条件</param>
        /// <param name="orderByFields">排序字段，如name asc,age desc</param>
        /// <returns>数据列表</returns>
        public async Task<List<TEntity>> Query(string where, string orderByFields)
        {
            return await db.Queryable<TEntity>().OrderByIF(!string.IsNullOrEmpty(orderByFields), orderByFields).WhereIF(!string.IsNullOrEmpty(where), where).ToListAsync();
        }

        /// <summary>
        /// 功能描述:查询前N条数据
        /// </summary>
        /// <param name="whereExpression">条件表达式</param>
        /// <param name="intTop">前N条</param>
        /// <param name="orderByFields">排序字段，如name asc,age desc</param>
        /// <returns>数据列表</returns>
        public async Task<List<TEntity>> Query(Expression<Func<TEntity, bool>> whereExpression, int intTop, string orderByFields)
        {
            return await db.Queryable<TEntity>().OrderByIF(!string.IsNullOrEmpty(orderByFields), orderByFields).WhereIF(whereExpression != null, whereExpression).Take(intTop).ToListAsync();
        }

        /// <summary>
        /// 功能描述:查询前N条数据
        /// </summary>
        /// <param name="where">条件</param>
        /// <param name="intTop">前N条</param>
        /// <param name="orderByFields">排序字段，如name asc,age desc</param>
        /// <returns>数据列表</returns>
        public async Task<List<TEntity>> Query(string where, int intTop, string orderByFields)
        {
            return await db.Queryable<TEntity>().OrderByIF(!string.IsNullOrEmpty(orderByFields), orderByFields).WhereIF(!string.IsNullOrEmpty(where), where).Take(intTop).ToListAsync();
        }

        /// <summary>
        /// 功能描述:分页查询
        /// </summary>
        /// <param name="whereExpression">条件表达式</param>
        /// <param name="pageIndex">页码（下标0）</param>
        /// <param name="pageSize">页大小</param>
        /// <param name="orderByFields">排序字段，如name asc,age desc</param>
        /// <returns>数据列表</returns>
        public async Task<List<TEntity>> Query(Expression<Func<TEntity, bool>> whereExpression, int pageIndex, int pageSize, string orderByFields)
        {
            return await db.Queryable<TEntity>().OrderByIF(!string.IsNullOrEmpty(orderByFields), orderByFields)
                 .WhereIF(whereExpression != null, whereExpression).ToPageListAsync(pageIndex, pageSize);
        }

        /// <summary>
        /// 功能描述:分页查询
        /// </summary>
        /// <param name="where">条件</param>
        /// <param name="pageIndex">页码（下标0）</param>
        /// <param name="pageSize">页大小</param>
        /// <param name="orderByFields">排序字段，如name asc,age desc</param>
        /// <returns>数据列表</returns>
        public async Task<List<TEntity>> Query(string where, int pageIndex, int pageSize, string orderByFields)
        {
            return await db.Queryable<TEntity>().OrderByIF(!string.IsNullOrEmpty(orderByFields), orderByFields)
                .WhereIF(!string.IsNullOrEmpty(where), where).ToPageListAsync(pageIndex, pageSize);
        }

        public async Task<TEntity> QueryById(object objId)
        {
            return await Task.Run(() => db.Queryable<TEntity>().InSingle(objId));
        }

        /// <summary>
        /// 功能描述:根据ID查询一条数据
        /// </summary>
        /// <param name="objId">id（必须指定主键特性 [SugarColumn(IsPrimaryKey=true)]），如果是联合主键，请使用Where条件</param>
        /// <param name="blnUseCache">是否使用缓存</param>
        /// <returns>数据实体</returns>
        public async Task<TEntity> QueryById(object objId, bool blnUseCache = false)
        {
            return await Task.Run(() => db.Queryable<TEntity>().WithCacheIF(blnUseCache).InSingle(objId));
        }

        /// <summary>
        /// 功能描述:根据ID查询数据
        /// </summary>
        /// <param name="lstIds">id列表（必须指定主键特性 [SugarColumn(IsPrimaryKey=true)]），如果是联合主键，请使用Where条件</param>
        /// <returns>数据实体列表</returns>
        public async Task<List<TEntity>> QueryByIDs(object[] lstIds)
        {
            return await Task.Run(() => db.Queryable<TEntity>().In(lstIds).ToList());
        }

        /// <summary>
        ///查询-多表查询
        /// </summary>
        /// <typeparam name="T">实体1</typeparam>
        /// <typeparam name="T2">实体2</typeparam>
        /// <typeparam name="T3">实体3</typeparam>
        /// <typeparam name="TResult">返回对象</typeparam>
        /// <param name="joinExpression">关联表达式 (join1,join2) => new object[] {JoinType.Left,join1.UserNo==join2.UserNo}</param>
        /// <param name="selectExpression">返回表达式 (s1, s2) => new { Id =s1.UserNo, Id1 = s2.UserNo}</param>
        /// <param name="whereLambda">查询表达式 (w1, w2) =>w1.UserNo == "")</param>
        public async Task<List<TResult>> QueryMuch<T, T2, T3, TResult>(Expression<Func<T, T2, T3, object[]>> joinExpression, Expression<Func<T, T2, T3, TResult>> selectExpression, Expression<Func<T, T2, T3, bool>> whereLambda = null) where T : class, new()
        {
            if (whereLambda == null)
            {
                return await db.Queryable(joinExpression).Select(selectExpression).ToListAsync();
            }
            return await db.Queryable(joinExpression).Where(whereLambda).Select(selectExpression).ToListAsync();
        }

        /// <summary>
        /// 两表联合查询-分页
        /// </summary>
        /// <typeparam name="T">实体1</typeparam>
        /// <typeparam name="T2">实体1</typeparam>
        /// <typeparam name="TResult">返回对象</typeparam>
        /// <param name="joinExpression">关联表达式</param>
        /// <param name="selectExpression">返回表达式</param>
        /// <param name="whereExpression">查询表达式</param>
        /// <param name="pageIndex">页码</param>
        /// <param name="pageSize">页大小</param>
        /// <param name="orderByFields">排序字段</param>
        /// <returns></returns>
        public async Task<PageModel<TEntity>> QueryPage(Expression<Func<TEntity, bool>> whereExpression, int pageIndex = 1, int pageSize = 20, string orderByFields = null)
        {
            RefAsync<int> totalCount = 0;
            var list = await db.Queryable<TEntity>()
             .OrderByIF(!string.IsNullOrEmpty(orderByFields), orderByFields)
             .WhereIF(whereExpression != null, whereExpression)
             .ToPageListAsync(pageIndex, pageSize, totalCount);

            return new PageModel<TEntity>(pageIndex, totalCount, pageSize, list);
        }

        /// <summary>
        /// 根据sql语句查询
        /// </summary>
        /// <param name="sql">完整的sql语句</param>
        /// <param name="parameters">参数</param>
        /// <returns>泛型集合</returns>
        public async Task<List<TEntity>> QuerySql(string sql, SugarParameter[] parameters = null)
        {
            return await db.Ado.SqlQueryAsync<TEntity>(sql, parameters);
        }

        /// <summary>
        /// 根据sql语句查询
        /// </summary>
        /// <param name="sql">完整的sql语句</param>
        /// <param name="parameters">参数</param>
        /// <returns>DataTable</returns>
        public async Task<DataTable> QueryTable(string sql, SugarParameter[] parameters = null)
        {
            return await db.Ado.GetDataTableAsync(sql, parameters);
        }

        /// <summary>
        /// 两表联合查询-分页
        /// </summary>
        /// <typeparam name="T">实体1</typeparam>
        /// <typeparam name="T2">实体1</typeparam>
        /// <typeparam name="TResult">返回对象</typeparam>
        /// <param name="joinExpression">关联表达式</param>
        /// <param name="selectExpression">返回表达式</param>
        /// <param name="whereExpression">查询表达式</param>
        /// <param name="pageIndex">页码</param>
        /// <param name="pageSize">页大小</param>
        /// <param name="orderByFields">排序字段</param>
        /// <returns></returns>
        public async Task<PageModel<TResult>> QueryTabsPage<T, T2, TResult>(Expression<Func<T, T2, object[]>> joinExpression, Expression<Func<T, T2, TResult>> selectExpression, Expression<Func<TResult, bool>> whereExpression, int pageIndex = 1, int pageSize = 20, string orderByFields = null)
        {
            RefAsync<int> totalCount = 0;
            var list = await db.Queryable<T, T2>(joinExpression)
             .Select(selectExpression)
             .OrderByIF(!string.IsNullOrEmpty(orderByFields), orderByFields)
             .WhereIF(whereExpression != null, whereExpression)
             .ToPageListAsync(pageIndex, pageSize, totalCount);
            return new PageModel<TResult>(pageIndex, totalCount, pageSize, list);
        }

        /// <summary>
        /// 两表联合查询-分页-分组
        /// </summary>
        /// <typeparam name="T">实体1</typeparam>
        /// <typeparam name="T2">实体1</typeparam>
        /// <typeparam name="TResult">返回对象</typeparam>
        /// <param name="joinExpression">关联表达式</param>
        /// <param name="selectExpression">返回表达式</param>
        /// <param name="whereExpression">查询表达式</param>
        /// <param name="groupExpression">group表达式</param>
        /// <param name="pageIndex">页码</param>
        /// <param name="pageSize">页大小</param>
        /// <param name="orderByFields">排序字段</param>
        /// <returns></returns>
        public async Task<PageModel<TResult>> QueryTabsPage<T, T2, TResult>(Expression<Func<T, T2, object[]>> joinExpression, Expression<Func<T, T2, TResult>> selectExpression, Expression<Func<TResult, bool>> whereExpression, Expression<Func<T, object>> groupExpression, int pageIndex = 1, int pageSize = 20, string orderByFields = null)
        {
            RefAsync<int> totalCount = 0;
            var list = await db.Queryable<T, T2>(joinExpression).GroupBy(groupExpression)
             .Select(selectExpression)
             .OrderByIF(!string.IsNullOrEmpty(orderByFields), orderByFields)
             .WhereIF(whereExpression != null, whereExpression)
             .ToPageListAsync(pageIndex, pageSize, totalCount);
            return new PageModel<TResult>(pageIndex, totalCount, pageSize, list);
        }

        public async Task<bool> Update(TEntity model)
        {
            return await db.Updateable(model).ExecuteCommandHasChangeAsync();
        }

        public async Task<bool> Update(TEntity entity, string where)
        {
            return await db.Updateable(entity).Where(where).ExecuteCommandHasChangeAsync();
        }

        public async Task<bool> Update(object operateAnonymousObjects)
        {
            return await db.Updateable<TEntity>(operateAnonymousObjects).ExecuteCommandAsync() > 0;
        }

        public async Task<bool> Update(TEntity entity, List<string> lstColumns = null, List<string> lstIgnoreColumns = null, string where = "")
        {
            IUpdateable<TEntity> up = db.Updateable(entity);
            if (lstIgnoreColumns != null && lstIgnoreColumns.Count > 0)
            {
                up = up.IgnoreColumns(lstIgnoreColumns.ToArray());
            }
            if (lstColumns != null && lstColumns.Count > 0)
            {
                up = up.UpdateColumns(lstColumns.ToArray());
            }
            if (!string.IsNullOrEmpty(where))
            {
                up = up.Where(where);
            }
            return await up.ExecuteCommandHasChangeAsync();
        }

        public async Task<bool> UpdateByWhere(TEntity entity, Action<TEntity> action)
        {
            IUpdateable<TEntity> up = db.Updateable(entity);
            return await up.ReSetValue(action).ExecuteCommandAsync() > 0;
        }
    }
}
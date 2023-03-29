/****************************************************************************************************************************************************************
 *
 * 版权所有: 2023  All Rights Reserved.
 * 文 件 名:  Result
 * 描    述:  数据结果对象
 * 创 建 者：  Ltf
 * 创建时间：  2023/3/29 15:37:03
 * 历史更新记录:
 * =============================================================================================
*修改标记
*修改时间：2023/3/29 15:37:03
*修改人： Ltf
*版本号： V1.0.0.0
*描述：
 * =============================================================================================

****************************************************************************************************************************************************************/

namespace CRCC.SystemTools
{
    public struct Result<T>
    {
        public Result(ResultStatus status, T value) : this(status, value, null)
        {
        }

        public Result(ResultStatus status, string message) : this(status, default(T), message)
        {
        }

        public Result(ResultStatus status) : this(status, default(T), null)
        {
        }

        public Result(string message) : this(ResultStatus.Error, default(T), message)
        {
        }

        public Result(ResultStatus status, T value, string message)
        {
            Status = status;
            Value = value;
            Message = message;
        }

        public Result(bool status, T value, string message = null) : this(status ? ResultStatus.Ok : ResultStatus.Error, value, message)
        {
        }

        /// <summary>
        /// 结果状态
        /// </summary>
        public ResultStatus Status { get; set; }

        /// <summary>
        /// 结果内容
        /// </summary>
        public T Value { get; set; }

        /// <summary>
        /// 结果消息
        /// </summary>
        public string Message { get; set; }

        public bool IsSuccess { get { return Status == ResultStatus.Ok; } }

        public static Result<T> SuccessNull = new Result<T>(ResultStatus.Ok, null);
    }

    /// <summary>
    /// 结果状态枚举
    /// </summary>
    public enum ResultStatus
    {
        Cancel = -1,
        Error = -2,
        No = 0,
        Ok = 1,
        Other = 2,
    }
}
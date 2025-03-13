using Microsoft.AspNetCore.Mvc.ApplicationModels;

namespace FeiNuo.AspNetCore
{
    /// <summary>
    /// 禁止自动校验数据模型
    /// </summary>
    [AttributeUsage(AttributeTargets.Method)]
    public class DisableModelValidationAttribute : Attribute, IActionModelConvention
    {
        private const string FilterName = "ModelStateInvalidFilterFactory";
        public void Apply(ActionModel action)
        {
            for (var i = 0; i < action.Filters.Count; i++)
            {
                if (action.Filters[i].GetType().Name == FilterName)
                {
                    action.Filters.RemoveAt(i);
                    break;
                }
            }
        }
    }
}

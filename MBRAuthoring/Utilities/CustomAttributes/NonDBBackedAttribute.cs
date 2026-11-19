using System;

namespace MBRAuthoring.Utilities.CustomAttributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class NonDBBackedAttribute : Attribute
    {
        public NonDBBackedAttribute()
        {

        }
    }
}

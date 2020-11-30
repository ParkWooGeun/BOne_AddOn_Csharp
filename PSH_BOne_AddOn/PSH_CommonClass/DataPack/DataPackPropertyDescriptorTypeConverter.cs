using System;
using System.Globalization;
using System.Reflection;
using System.ComponentModel;
using System.ComponentModel.Design.Serialization;

namespace PSH_BOne_AddOn.Database.Pack
{
    public class DataPackPropertyDescriptorTypeConverter : TypeConverter
    {
        /// <summary>
        /// TypeConverter 의 CanConvertTo 구현
        /// </summary>
        /// <param name="context"></param>
        /// <param name="destinationType"></param>
        /// <returns></returns>
        public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
        {
            if (destinationType == typeof(InstanceDescriptor))
                return true;

            return base.CanConvertTo(context, destinationType);
        }

        /// <summary>
        /// TypeConverter 의 ConvertTo 구현
        /// </summary>
        /// <param name="context"></param>
        /// <param name="culture"></param>
        /// <param name="value"></param>
        /// <param name="destinationType"></param>
        /// <returns></returns>
        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType)
        {
            if (destinationType == typeof(InstanceDescriptor) && value is DataPackPropertyDescriptor)
            {
                DataPackPropertyDescriptor pd = (DataPackPropertyDescriptor)value;

                ConstructorInfo ctor = typeof(DataPackPropertyDescriptor).GetConstructor(new Type[] {typeof(string), typeof(Type)});

                if (ctor != null)
                {
                    return new InstanceDescriptor(ctor, new object[] {pd.Name, pd.PropertyType});
                }
            }
            return base.ConvertTo(context, culture, value, destinationType);
        }
    }
}

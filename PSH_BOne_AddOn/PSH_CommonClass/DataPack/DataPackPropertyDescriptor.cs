using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.IO;

namespace PSH_BOne_AddOn.Helper.Data
{
    [
        TypeConverter(typeof(PSH_BOne_AddOn.Helper.Data.DataPackPropertyDescriptorTypeConverter))
    ]
    public class DataPackPropertyDescriptor : PropertyDescriptor
    {
        /// <summary>
        /// 
        /// </summary>
        private Type _propertyType = null;

        /// <summary>
        /// 
        /// </summary>
        private object _value;

        /// <summary>
        /// 
        /// </summary>
        private static CategoryAttribute _defaultAttr = null;

        /// <summary>
        /// 
        /// </summary>
        private AttributeCollection _attrs = null;

        /// <summary>
        /// 
        /// </summary>
        static DataPackPropertyDescriptor()
        {
            _defaultAttr = new CategoryAttribute(DataPack.CATEGORY_CUSTOMPROPERTY);
        }

        /// <summary>
        /// DataPackPropertyDescriptor 생성자
        /// </summary>
        /// <param name="name"></param>
        /// <param name="propertyType"></param>
        public DataPackPropertyDescriptor(string name, Type propertyType) : this(name, propertyType, null)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="propertyType"></param>
        /// <param name="attributes"></param>
        public DataPackPropertyDescriptor(string name, Type propertyType, Attribute[] attributes) : base(name, attributes)
        {
            _propertyType = propertyType;
            if (propertyType.Name == "String")
            {
                _value = string.Empty;
            }
            else
            {
                _value = CreateInstance(propertyType);
            }
        }

        /// <summary>
        /// PropertyDescriptor 의 SetValue 구현
        /// </summary>
        /// <param name="component"></param>
        /// <param name="value"></param>
        public override void SetValue(object component, object value)
        {
            IComponent comp = component as IComponent;
            if (comp != null && comp.Site != null && comp.Site.DesignMode == true)
            {
                IDesignerHost idh = comp.Site.GetService(typeof(IDesignerHost)) as IDesignerHost;

                IComponentChangeService icc = comp.Site.GetService(typeof(IComponentChangeService)) as IComponentChangeService;

                if (idh != null && icc != null)
                {
                    DesignerTransaction tx = null;
                    try
                    {
                        tx = idh.CreateTransaction("DataPackPropertyDescriptor SetValue");

                        icc.OnComponentChanging(comp, this);

                        object oldValue = _value;
                        _value = value;

                        icc.OnComponentChanged(comp, this, oldValue, value);

                        tx.Commit();
                    }
                    catch (System.Exception ex)
                    {
                        if (tx != null)
                        {
                            tx.Cancel();
                        }

                        StreamWriter sw = File.AppendText("c:\\designtimelog.txt");
                        sw.Write(ex.ToString());
                        sw.Close();

                        throw new ArgumentException("DataPackPropertyDescriptor.SetValue", ex);
                    }
                }
            }
            else
            {
                _value = value;
            }
        }

        /// <summary>
        /// PropertyDescriptor 의 GetValue 구현
        /// </summary>
        /// <param name="component"></param>
        /// <returns></returns>
        public override object GetValue(object component)
        {
            return _value;
        }

        /// <value>
        /// PropertyDescriptor 의 IsReadOnly 구현
        /// </value>
        public override bool IsReadOnly
        {
            get
            {
                return false;
            }
        }

        /// <value>
        /// PropertyDescriptor 의 ComponentType 구현
        /// </value>
        public override System.Type ComponentType
        {
            get
            {
                return typeof(DataPack);
            }
        }

        /// <value>
        /// PropertyDescriptor 의 PropertyType 구현
        /// </value>
        public override System.Type PropertyType
        {
            get
            {
                return _propertyType;
            }
        }

        /// <summary>
        /// PropertyDescriptor 의 CanResetValue 구현
        /// </summary>
        /// <param name="component"></param>
        /// <returns></returns>
        public override bool CanResetValue(object component)
        {
            return true;
        }

        /// <summary>
        /// PropertyDescriptor 의 ResetValue 구현
        /// </summary>
        /// <param name="component"></param>
        public override void ResetValue(object component)
        {
            if (this.PropertyType.Name == "String")
            {
                _value = string.Empty;
            }
            else
            {
                _value = CreateInstance(this.PropertyType);
            }
        }

        /// <summary>
        /// PropertyDescriptor 의 ShouldSerializeValue 구현
        /// </summary>
        /// <param name="component"></param>
        /// <returns></returns>
        public override bool ShouldSerializeValue(object component)
        {
            return true;
        }

        /// <summary>
        /// MemberDescriptor 의 Attributes 구현
        /// </summary>
        public override System.ComponentModel.AttributeCollection Attributes
        {
            get
            {
                if (_attrs == null)
                {
                    // 원래의 특성배열을  새로운 특성배열로 복사하고 카테고리 특성을
                    // 하나 추가한다.
                    Attribute[] attrs = new Attribute[base.AttributeArray.Length + 2];
                    base.AttributeArray.CopyTo(attrs, 0);

                    attrs[base.AttributeArray.Length] = new CategoryAttribute(DataPack.CATEGORY_CUSTOMPROPERTY);
                    attrs[base.AttributeArray.Length + 1] = new DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Hidden);

                    AttributeArray = attrs;

                    _attrs = new AttributeCollection(attrs);
                }
                return _attrs;
            }
        }
    }
}


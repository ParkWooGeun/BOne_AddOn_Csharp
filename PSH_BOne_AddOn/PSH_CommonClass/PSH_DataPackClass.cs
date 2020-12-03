namespace PSH_BOne_AddOn.DataPack
{
    public class PSH_DataPackClass
    {
        public PSH_DataPackClass(object code, object value)
        {
            this.Code = code;
            this.Value = value;
        }

        public PSH_DataPackClass(object code, object value, object type)
        {
            this.Code = code;
            this.Value = value;
            this.Type = type;
        }

        public object Code
        {
            get;
            set;
        }

        public object Value
        {
            get;
            set;
        }

        public object Type
        {
            get;
            set;
        }
    }
}

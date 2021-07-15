using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.ComponentModel;
using System.Collections;
using System.Runtime.Serialization;

namespace PSH_BOne_AddOn.Database.Pack
{
    /// <remarks>
    /// Tier, Biz 간에 DataParameter를 한번에 묶어서 보내기 위한 HelperClass.
    /// 이는 타시스템 연동을 위한 Serialize 지원
    /// DataSet,DataTable,DataRow 와 DataPack 간에 자료 공유
    /// DataPack 의 내용으로부터 DBParameter[] 생성등 다양한 Helper 메서드들이 존재한다.
    /// </remarks>
    [
        Serializable()
    ]
    public sealed class DataPack : MarshalByValueComponent, IComponent, ITypedList, IList, ICustomTypeDescriptor, ISerializable
    {
        /// <summary>
        /// 사용자속성
        /// </summary>
        public const string CATEGORY_CUSTOMPROPERTY = "사용자속성";

        /// <summary>
        /// 실제속성
        /// </summary>
        public const string CATEGORY_REALPROPERTY = "실제속성";

        /// <summary>
        /// 
        /// </summary>
        private static CategoryAttribute _customPropertyAttr = null;

        private ArrayList _rows = null;
        private PropertyDescriptorCollection _properties = null;

        /// <summary>
        /// 
        /// </summary>
        static DataPack()
        {
            _customPropertyAttr = new CategoryAttribute(CATEGORY_CUSTOMPROPERTY);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="container"></param>
        public DataPack(IContainer container)
        {
            InitClass();

            container.Add(this);
        }

        /// <summary>
        /// DataPack를 초기화 시키는 메서드를 호출하는 클래스 생성자
        /// </summary>
        public DataPack()
        {
            InitClass();
        }

        /// <summary>
        /// Serialization 된 DataPack 으로부터 DataPack 재구성 
        /// </summary>
        /// <param name="info">데이터로 채워진 SerializationInfo</param>
        /// <param name="context">이 serialization에 대한 대상</param>
        public DataPack(SerializationInfo info, StreamingContext context)
        {
            InitClass();

            SerializationInfoEnumerator enumerator = info.GetEnumerator();
            object value;
            while (enumerator.MoveNext() == true)
            {
                value = enumerator.Value.ToString().Trim();
                AddProperty(enumerator.Name, value.GetType(), value);
            }
        }

        /// <summary>
        /// DataPack를 초기화 시키는 메서드 
        /// </summary>
        private void InitClass()
        {
            _rows = new ArrayList();
            _rows.Add(this);

            _properties = new PropertyDescriptorCollection(null);
        }


        #region Implementation of ISerializable
        /// <summary>
        /// DataPack를 Serialize를 하기 위해서 구현된 메서드
        /// </summary>
        /// <param name="info">데이터로 채울 SerializationInfo</param>
        /// <param name="context">이 serialization에 대한 대상</param>
        public void GetObjectData(System.Runtime.Serialization.SerializationInfo info, System.Runtime.Serialization.StreamingContext context)
        {
            CategoryAttribute customPropertyAttr = new CategoryAttribute(CATEGORY_CUSTOMPROPERTY);

            foreach (PropertyDescriptor pd in _properties)
            {
                if (pd.Attributes.Contains(customPropertyAttr) == true)
                {
                    info.AddValue(pd.Name, pd.GetValue(this), pd.PropertyType);
                }
            }
        }
        #endregion

        #region Implementation of ITypedList
        /// <summary>
        /// 
        /// </summary>
        /// <param name="listAccessors"></param>
        /// <returns></returns>
        PropertyDescriptorCollection ITypedList.GetItemProperties(PropertyDescriptor[] listAccessors)
        {
            ICustomTypeDescriptor ictd = this as ICustomTypeDescriptor;
            Attribute[] attrs = { _customPropertyAttr };

            return ictd.GetProperties(attrs);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listAccessors"></param>
        /// <returns></returns>
        string ITypedList.GetListName(PropertyDescriptor[] listAccessors)
        {
            return Site.Name;
        }
        #endregion

        #region Implementation of IList
        void IList.RemoveAt(int index)
        {
        }

        void IList.Insert(int index, object value)
        {

        }

        void IList.Remove(object value)
        {

        }

        bool IList.Contains(object value)
        {
            return _rows.Contains(value);
        }

        void IList.Clear()
        {

        }

        int IList.IndexOf(object value)
        {
            return 0;
        }

        int IList.Add(object value)
        {
            return 0;
        }

        bool IList.IsReadOnly
        {
            get
            {
                return true;
            }
        }

        object IList.this[int index]
        {
            get
            {
                return _rows[index];
            }
            set
            {
            }
        }

        bool IList.IsFixedSize
        {
            get
            {
                return true;
            }
        }
        #endregion

        #region Implementation of ICollection
        void ICollection.CopyTo(System.Array array, int index)
        {
            _rows.CopyTo(array, index);
        }

        bool ICollection.IsSynchronized
        {
            get
            {
                return _rows.IsSynchronized;
            }
        }

        int ICollection.Count
        {
            get
            {
                return _rows.Count;
            }
        }

        object ICollection.SyncRoot
        {
            get
            {
                return _rows.SyncRoot;
            }
        }
        #endregion

        #region Implementation of IEnumerable
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _rows.GetEnumerator();
        }
        #endregion

        #region Implementation of ICustomTypeDescriptor
        TypeConverter ICustomTypeDescriptor.GetConverter()
        {
            return null;
        }

        EventDescriptorCollection ICustomTypeDescriptor.GetEvents(System.Attribute[] attributes)
        {
            return ((ICustomTypeDescriptor)this).GetEvents();
        }

        EventDescriptorCollection ICustomTypeDescriptor.GetEvents()
        {
            return EventDescriptorCollection.Empty;
        }

        string ICustomTypeDescriptor.GetComponentName()
        {
            if (this.Site != null)
            {
                return Site.Name;
            }
            return null;
        }

        object ICustomTypeDescriptor.GetPropertyOwner(PropertyDescriptor pd)
        {
            return this;
        }

        AttributeCollection ICustomTypeDescriptor.GetAttributes()
        {
            return TypeDescriptor.GetAttributes(this, true);
        }

        PropertyDescriptorCollection ICustomTypeDescriptor.GetProperties(System.Attribute[] attributes)
        {
            PropertyDescriptorCollection pdc = new PropertyDescriptorCollection(null);

            foreach (PropertyDescriptor pd in ((ICustomTypeDescriptor)this).GetProperties())
            {
                if (pd.Attributes.Contains(attributes) == true)
                {
                    pdc.Add(pd);
                }
            }
            return pdc;
        }

        PropertyDescriptorCollection ICustomTypeDescriptor.GetProperties()
        {
            return _properties;
        }

        object ICustomTypeDescriptor.GetEditor(System.Type editorBaseType)
        {
            return null;
        }

        PropertyDescriptor ICustomTypeDescriptor.GetDefaultProperty()
        {
            return null;
        }

        EventDescriptor ICustomTypeDescriptor.GetDefaultEvent()
        {
            return null;
        }

        string ICustomTypeDescriptor.GetClassName()
        {
            return "DataPack";
        }
        #endregion

        #region Implementation of IComponent
        ISite IComponent.Site
        {
            get
            {
                return base.Site;
            }
            set
            {
                base.Site = value;
                if (_properties.Count < 1)
                {
                    PropertyDescriptorCollection realProperties = TypeDescriptor.GetProperties(this, true);

                    foreach (PropertyDescriptor item in realProperties)
                    {
                        _properties.Add(item);
                    }
                }
            }
        }
        #endregion

        //    [
        //    Browsable(false),
        //    Category(CATEGORY_REALPROPERTY)
        //    ]
        //    public ICollection Rows
        //    {
        //      get
        //      {
        //        return _rows as ICollection;
        //      }
        //    }

        /// <value>
        /// DataPack 내에 있는 PropertyDescriptorCollection 객체를 직접 얻어낼수 있다.
        /// </value>
        [
            Browsable(true),
            Category(CATEGORY_REALPROPERTY),
            NotifyParentProperty(true),
            RefreshProperties(RefreshProperties.All),
            //Editor(typeof(DataPackPropertyCollectionEditor),typeof(UITypeEditor)),
            DesignerSerializationVisibility(DesignerSerializationVisibility.Content),
        ]
        public PropertyDescriptorCollection Properties
        {
            get
            {
                return _properties;
            }
        }

        /// <value>
        /// DataPack 내에 지정된 이름을 통해서 프로퍼티 값을 얻어내거나 할당한다.
        /// 존재하지 않을때 ArgumentException가 발생한다.
        /// </value>
        public object this[string name]
        {
            get
            {
                DataPackPropertyDescriptor pd = _properties.Find(name, false) as DataPackPropertyDescriptor;

                if (pd != null)
                {
                    return pd.GetValue(this);
                }
                else
                {
                    throw new ArgumentException("DataPack 내에 " + name + "이라는 property가 존재하지 않습니다");
                }
            }
            set
            {
                DataPackPropertyDescriptor pd = _properties.Find(name, false) as DataPackPropertyDescriptor;

                if (pd != null)
                {
                    pd.SetValue(this, value);
                }
                else
                {
                    throw new ArgumentException("DataPack 내에 " + name + "이라는 property가 존재하지 않습니다");
                }
            }
        }

        /// <value>
        /// DataPack에 있는 프로퍼티의 값들을 ICollection 인터페이스를 통해서 얻어낼수 있다.
        /// </value>
        [
            Browsable(false),
            Category(CATEGORY_REALPROPERTY),
        ]
        public ICollection Keys
        {
            get
            {
                CategoryAttribute customPropertyAttr = new CategoryAttribute(CATEGORY_CUSTOMPROPERTY);

                ArrayList arrayList = new ArrayList();

                foreach (PropertyDescriptor pd in _properties)
                {
                    if (pd.Attributes.Contains(customPropertyAttr) == true)
                    {
                        arrayList.Add(pd.Name);
                    }
                }

                ICollection col = arrayList as ICollection;

                return col;
            }
        }

        /// <value>
        /// DataPack에 있는 프로퍼티들을 ICollection 인터페이스를 통해서 얻어낼수 있다.
        /// </value>
        [
            Browsable(false),
            Category(CATEGORY_REALPROPERTY),
        ]
        public ICollection Values
        {
            get
            {
                CategoryAttribute customPropertyAttr = new CategoryAttribute(CATEGORY_CUSTOMPROPERTY);

                ArrayList arrayList = new ArrayList();

                foreach (PropertyDescriptor pd in _properties)
                {
                    if (pd.Attributes.Contains(customPropertyAttr) == true)
                    {
                        arrayList.Add(pd.GetValue(this));
                    }
                }

                ICollection col = arrayList as ICollection;

                return col;
            }
        }

        /// <summary>
        /// DataPack 내에 프로퍼티가 있는지 조사하는 메서드
        /// </summary>
        /// <param name="propertyName">조회하고조 하는 프로퍼티 이름</param>
        /// <returns>존재유무(true/false)</returns>
        public bool Contains(string propertyName)
        {
            DataPackPropertyDescriptor pd = _properties.Find(propertyName, false)
                as DataPackPropertyDescriptor;

            return (pd != null) ? true : false;
        }

        #region Add/Remove/ResetAllValues
        /// <summary>
        /// DataPack에 새로운 프로퍼티를 추가한다. (초기값 지정가능)
        /// </summary>
        /// <param name="name">추가할 프로퍼티 이름</param>
        /// <param name="type">프로퍼티 타입</param>
        /// <param name="initialValue">초기값</param>
        public void AddProperty(string name, Type type, object initialValue)
        {
            DataPackPropertyDescriptor pd = new DataPackPropertyDescriptor(name, type);

            _properties.Add(pd);
            pd.SetValue(this, initialValue);
        }

        /// <summary>
        /// DataPack에 새로운 프로퍼티를 추가한다.
        /// </summary>
        /// <param name="name">추가할 프로퍼티 이름</param>
        /// <param name="type">프로퍼티 타입</param>
        public void AddProperty(string name, Type type)
        {
            DataPackPropertyDescriptor pd = new DataPackPropertyDescriptor(name, type);

            _properties.Add(pd);
        }

        /// <summary>
        /// DataPack에서 지정된 이름의 프로퍼티를 삭제한다.
        /// </summary>
        /// <param name="name">삭제할 프로퍼티 이름</param>
        public void RemoveProperty(string name)
        {
            DataPackPropertyDescriptor pd = _properties.Find(name, false) as DataPackPropertyDescriptor;

            if (pd != null)
            {
                _properties.Remove(pd);
            }
            else
            {
                throw new ArgumentException();
            }
        }

        /// <summary>
        /// DataPack에 있는 모든 프로퍼티들을 삭제시킨다.
        /// </summary>
        public void ResetAllValues()
        {
            ITypedList itl = this as ITypedList;
            foreach (PropertyDescriptor pd in itl.GetItemProperties(null))
            {
                if (pd.CanResetValue(this) == true)
                {
                    pd.ResetValue(this);
                }
                else
                {
                    pd.SetValue(this, null);
                }
            }
        }
        #endregion

        #region import FROM
        /// <summary>
        /// DataRow 로부터 DataPack 내용을 새롭게 생성한다.
        /// </summary>
        /// <param name="dataRow">DataRow 객체</param>
        /// <returns>DataPack 객체</returns>
        public static DataPack FromDataRow(DataRow dataRow)
        {
            DataPack dataPack = new DataPack();

            foreach (DataColumn col in dataRow.Table.Columns)
            {
                string colName = col.ColumnName;
                dataPack.AddProperty(col.ColumnName, col.DataType, dataRow[col]);
            }

            return dataPack;
        }

        /// <summary>
        /// DataTable 에는 DataRow가 Array 형태로 들어 있다.
        /// 있을 DataPack[] 형태로 만들어낸다.
        /// </summary>
        /// <param name="table">DataTable 객체</param>
        /// <returns>DataPack[] 객체</returns>
        public static DataPack[] FromDataTable(DataTable table)
        {
            if (table == null)
            {
                throw new ArgumentNullException();
            }

            if (table.Rows.Count < 1)
            {
                return null;
            }

            DataPack[] dataPacks = new DataPack[table.Rows.Count];

            for (int i = 0; i < table.Rows.Count; i++)
            {
                dataPacks[i] = DataPack.FromDataRow(table.Rows[i]);
            }

            return dataPacks;
        }

        /// <summary>
        /// DataSet 의 첫번째 DataTable 값을 DataPack[]로 만들어낸다.
        /// </summary>
        /// <param name="ds">DataSet 객체</param>
        /// <returns>DataPack[] 객체</returns>
        public static DataPack[] FromDataSet(DataSet ds)
        {
            if (ds == null)
            {
                throw new ArgumentNullException();
            }

            if (ds.Tables.Count < 1 || ds.Tables[0].Rows.Count < 1)
            {
                return null;
            }

            return DataPack.FromDataTable(ds.Tables[0]);
        }

        #endregion

        #region Export To 

        /// <summary>
        /// DataPack 내용을 기존에 DataRow 로 추가한다.
        /// </summary>
        /// <param name="dataRow">DataRow 객체</param>
        public void ToDataRow(DataRow dataRow)
        {
            if (dataRow == null)
            {
                throw new ArgumentNullException();
            }

            ITypedList itl = this as ITypedList;
            foreach (PropertyDescriptor pd in itl.GetItemProperties(null))
            {
                if (dataRow.Table.Columns.Contains(pd.Name))
                {
                    dataRow[pd.Name] = pd.GetValue(this);
                }
            }
        }

        /// <summary>
        /// DataPack 내용으로 부터 새로운 DataRow 를 만들어낸다.
        /// </summary>
        /// <returns></returns>
        public DataRow ToDataRow()
        {
            DataTable table = new DataTable();
            DataRow dataRow = table.NewRow();

            ITypedList itl = this as ITypedList;
            foreach (PropertyDescriptor pd in itl.GetItemProperties(null))
            {
                table.Columns.Add(pd.Name, pd.PropertyType);
                dataRow[pd.Name] = pd.GetValue(this);
            }

            return dataRow;
        }

        /// <summary>
        /// DataPack 내용으로 부터 XML 문자열을 생성
        /// </summary>
        /// <returns>xml 문자열</returns>
        public string ToXMLString()
        {
            DataSet ds = new DataSet();
            DataTable table = new DataTable();
            DataRow dataRow = table.NewRow();

            ITypedList itl = this as ITypedList;
            foreach (PropertyDescriptor pd in itl.GetItemProperties(null))
            {
                table.Columns.Add(pd.Name, pd.PropertyType);
                dataRow[pd.Name] = pd.GetValue(this);
            }
            table.Rows.Add(dataRow);
            ds.Tables.Add(table);

            return ds.GetXml();
        }

        #endregion

        #region ToSqlParameters
        /// <summary>
        /// DataPack 을 통해서 SqlParameters[] 를 만들어 낸다.
        /// </summary>
        /// <param name="useAtSign">SP 에서 하는 '@' 문자열을 추가여부 지정</param>
        /// <returns>SqlParameter[] 객체</returns>
        public SqlParameter[] ToSqlParameters(bool useAtSign)
        {
            ITypedList itl = this as ITypedList;
            PropertyDescriptorCollection pdc = itl.GetItemProperties(null);

            int count = pdc.Count;

            if (count < 1)
            {
                throw new InvalidOperationException();
            }

            SqlParameter[] parameters
                = new SqlParameter[count];

            int index = 0;
            foreach (PropertyDescriptor pd in pdc)
            {
                string keyName = string.Empty;
                if (useAtSign == true)
                {
                    keyName += ("@" + pd.Name);
                }

                parameters[index++] = new SqlParameter(keyName, pd.GetValue(this));
            }

            return parameters;
        }

        /// <summary>
        /// DataPack 을 통해서 SqlParameters[] 를 만들어 낸다.
        /// SP에서 사용되는 변수처럼 '@' 가 추가된SqlParameters 를 만든다.
        /// </summary>
        /// <returns>SqlParameter[] 객체</returns>
        public SqlParameter[] ToSqlParameters()
        {
            return ToSqlParameters(true);
        }

        /// <summary>
        /// 기존에 이미 생성된 SqlParameter[]에 DataPack의 내용을 SqlParameter[]로
        /// 만들고 이를 추가시킨다.
        /// </summary>
        /// <param name="parameters">기존 SqlParameter[]</param>
        public void ToSqlParameters(SqlParameter[] parameters)
        {
            ITypedList itl = this as ITypedList;
            PropertyDescriptorCollection pdc = itl.GetItemProperties(null);

            int count = pdc.Count;

            if (count < 1)
            {
                throw new InvalidOperationException();
            }
            if (parameters.Length != count)
            {
                throw new ArgumentException();
            }

            int index = 0;
            foreach (PropertyDescriptor pd in pdc)
            {
                parameters[index++].Value = pd.GetValue(this);
            }
        }
        #endregion

        #region ToOleDbParameters

        /// <summary>
        /// DataPack 을 통해서 OleDbParameter[] 를 만들어 낸다.
        /// </summary>
        /// <param name="useAtSign">SP 에서 하는 '?' 문자열을 추가여부 지정</param>
        /// <returns>OleDbParameter[] 객체</returns>
        public OleDbParameter[] ToOleDbParameters(bool useAtSign)
        {
            ITypedList itl = this as ITypedList;
            PropertyDescriptorCollection pdc = itl.GetItemProperties(null);

            int count = pdc.Count;

            if (count < 1)
            {
                throw new InvalidOperationException();
            }

            OleDbParameter[] parameters = new OleDbParameter[count];

            int index = 0;
            foreach (PropertyDescriptor pd in pdc)
            {
                string keyName = string.Empty;
                if (useAtSign == true)
                {
                    keyName += ("@" + pd.Name);
                }
                parameters[index++] = new OleDbParameter(keyName, pd.GetValue(this));
            }

            return parameters;
        }

        /// <summary>
        /// DataPack 을 통해서 SqlParameters[] 를 만들어 낸다.
        /// SP에서 사용되는 변수처럼 '@' 가 추가된SqlParameters 를 만든다.
        /// </summary>
        /// <returns>SqlParameter[] 객체</returns>
        public OleDbParameter[] ToOleDbParameters()
        {
            return ToOleDbParameters(true);
        }

        /// <summary>
        /// 기존에 이미 생성된 OleDbParameter[]에 DataPack의 내용을 OleDbParameter[]로
        /// 만들고 이를 추가시킨다.
        /// </summary>
        /// <param name="parameters">기존 OleDbParameter[]</param>
        public void ToOleDbParameters(OleDbParameter[] parameters)
        {
            ITypedList itl = this as ITypedList;
            PropertyDescriptorCollection pdc = itl.GetItemProperties(null);

            int count = pdc.Count;

            if (count < 1)
            {
                throw new InvalidOperationException();
            }
            if (parameters.Length != count)
            {
                throw new ArgumentException();
            }

            int index = 0;
            foreach (PropertyDescriptor pd in pdc)
            {
                parameters[index++].Value = pd.GetValue(this);
            }
        }
        #endregion

        #region ToOracleParameters : ORACLE_ENABLE 정의에 의해서 활성화
#if ORACLE_ENABLE
		public OracleParameter[] ToOracleParameters(bool useColonSign)
		{
			ITypedList itl = this as ITypedList;
			PropertyDescriptorCollection pdc = itl.GetItemProperties(null);

			int count = pdc.Count;

			if ( count < 1 )
			{
				throw new InvalidOperationException();
			}

			OracleParameter[] parameters 
				= new OracleParameter[count];

			int index = 0;
			foreach( PropertyDescriptor pd in pdc )
			{
				string keyName = string.Empty;
				if ( useColonSign == true)
				{
					keyName += (":" + pd.Name);
				}
				parameters[index++]
					= new OracleParameter(keyName,pd.GetValue(this));
			}

			return parameters;
		}

		public OracleParameter[] ToOracleParameters()
		{
			return ToOracleParameters(true);
		}

		// 쿼리에 나타난 순서에 맞게 parameter 컬렉션을 구성해서 반환
		public OracleParameter[] ToOracleParameters( string query )
		{
			ITypedList itl = this as ITypedList;
			PropertyDescriptorCollection pdc = itl.GetItemProperties(null);

			int count = pdc.Count;

			if ( count < 1 )
			{
				throw new InvalidOperationException();
			}

			// query 상에서 ":" 으로 되어 있는 파라미터 개수만큼 Parameter[] 를 생성
			int nCount = StringEx.GetCountWords( query, ":" );

			ArrayList arrParam = new ArrayList();
			Hashtable hashPos = new Hashtable();

			foreach( PropertyDescriptor pd in pdc )
			{
				string keyName = pd.Name;
				
				// DataPack 에는 있으나, 쿼리에는 없는 컬럼은 parameter 구성에서 제외
				int nPos = query.IndexOf( ":" + keyName );
				if ( nPos == -1 )
				{
					continue;
				}

				// DataPack 에 같은이름의 컬럼이 중복되어 있는 경우, parameter 구성에는 하나만 사용
				object prevValue = hashPos[ nPos ];
				if ( prevValue != null ) 
				{
					continue;
				}

				OracleParamHelper param = new OracleParamHelper( nPos, keyName, pd.GetValue( this ) );
				arrParam.Add( param );
				hashPos.Add( nPos, string.Empty );
			}

			OracleParamHelper[] paramArr = (OracleParamHelper[])arrParam.ToArray( typeof( OracleParamHelper ) );

			Array.Sort( paramArr );

			OracleParameter[] parameters = new OracleParameter[ paramArr.Length ];

			int index = 0;
			foreach ( OracleParamHelper item in paramArr )
			{
				parameters[ index ++ ] = new OracleParameter( item.Name, item.Value );
			}

			return parameters;
		}

		public void ToOracleParameters(OracleParameter[] parameters)
		{
			ITypedList itl = this as ITypedList;
			PropertyDescriptorCollection pdc = itl.GetItemProperties(null);

			int count = pdc.Count;

			if ( count < 1 )
			{
				throw new InvalidOperationException();
			}
			if ( parameters.Length != count )
			{
				throw new ArgumentException();
			}

			int index = 0;
			foreach( PropertyDescriptor pd in pdc )
			{
				parameters[index++].Value = pd.GetValue(this);
			}
		}
#endif
        #endregion

        #region Merge
        /// <summary>
        /// 이미 존재하는 DataPack 내용에 DataRow 값을 추가한다.
        /// </summary>
        /// <param name="dataRow">추가하고자 하는 DataRow</param>
        public void Merge(DataRow dataRow)
        {
            ITypedList itl = this as ITypedList;
            foreach (PropertyDescriptor pd in itl.GetItemProperties(null))
            {
                pd.SetValue(this, dataRow[pd.Name]);
            }
        }

        /// <summary>
        /// 이미 존재하는 DataPack 내용에  또다른 DataPack를 추가한다.
        /// </summary>
        /// <param name="dataPack">새롭게 추가할 DataPack</param>
        public void Merge(DataPack dataPack)
        {
            ITypedList itl = this as ITypedList;
            foreach (PropertyDescriptor pd in itl.GetItemProperties(null))
            {
                pd.SetValue(this, dataPack[pd.Name]);
            }
        }
        #endregion

        #region Copy
        /// <summary>
        /// 이미 존재하는 DataPack으로부터 값을 복사해서 새로운 DataPack을 생성한다.
        /// </summary>
        /// <returns>기존값을 복사한 새로운 DataPack 객체</returns>
        public DataPack Copy()
        {
            DataPack dataPack = new DataPack();

            ITypedList itl = this as ITypedList;
            foreach (PropertyDescriptor pd in itl.GetItemProperties(null))
            {
                dataPack.AddProperty(pd.Name, pd.PropertyType, pd.GetValue(this));
            }

            return dataPack;
        }
        #endregion
    }
}

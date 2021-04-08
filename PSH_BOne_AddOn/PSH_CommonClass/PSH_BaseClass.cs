namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 기초클래스 : 모든 화면 클래스가 상속 받아 사용, 화면클래스가 기본적으로 사용할 속성 및 메소드 정의
    /// </summary>
    public class PSH_BaseClass
    {
        public SAPbouiCOM.Form oForm;

        public virtual void LoadForm()
        {
        }

        public virtual void LoadForm(string oFormDocEntry01)
        {
        }

        public virtual void VirtualFormItemEnabled()
        {
        }

        public virtual void LoadForm(string FromDate, string ToDate, string BPLID, string StdDt, string TabID)
        {
        }

        ///// <summary>
        ///// PS_SM010 Type 경우 사용됨
        ///// </summary>
        ///// <param name="Form"></param>
        ///// <param name="ItemUID"></param>
        ///// <param name="ColUID"></param>
        ///// <param name="ColRow"></param>
        //public virtual void LoadForm(SAPbouiCOM.Form Form, string ItemUID, string ColUID, int ColRow)
        //{
        //}

        public virtual void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        {
        }

        public virtual void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
        }

        public virtual void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
        }

        public virtual void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        {
        }

    }
}

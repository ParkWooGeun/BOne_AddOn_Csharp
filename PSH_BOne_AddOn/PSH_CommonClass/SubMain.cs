namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 최초 실행 클래스
    /// </summary>
	static class SubMain
    {
        /// <summary>
        /// 최초실행 Method
        /// </summary>
        static void Main()
        {
            PSH_MainClass Application = new PSH_MainClass();

            while (MessageAPIs.GetMessage(ref MessageAPIs.Msg_Renamed, 0, 0, 0))
            {
                MessageAPIs.TranslateMessage(ref MessageAPIs.Msg_Renamed);
                MessageAPIs.DispatchMessage(ref MessageAPIs.Msg_Renamed);
                System.Windows.Forms.Application.DoEvents();
            }

            System.Windows.Forms.Application.Run();
        }

        /// <summary>
        /// 폼객체 추가
        /// </summary>
        /// <param name="cObject"></param>
        /// <param name="oFormUid"></param>
        /// <param name="oFormTypeEx"></param>
        public static void Add_Forms(object cObject, string oFormUid, object oFormTypeEx = null)
        {
            PSH_Globals.ClassList.Add(cObject, oFormUid);
            PSH_Globals.FormTotalCount = PSH_Globals.FormTotalCount + 1;
            PSH_Globals.FormCurrentCount = PSH_Globals.FormCurrentCount + 1;
            PSH_Globals.FormTypeList.Add(oFormTypeEx, PSH_Globals.FormTypeListCount.ToString());
            PSH_Globals.FormTypeListCount = PSH_Globals.FormTypeListCount + 1;
        }

        /// <summary>
        /// 폼객체 제거
        /// </summary>
        /// <param name="oFormUniqueID"></param>
        public static void Remove_Forms(string oFormUniqueID)
        {
            object oTempClass = null;

            oTempClass = PSH_Globals.ClassList[oFormUniqueID];
            PSH_Globals.ClassList.Remove(oFormUniqueID);
            PSH_Globals.FormCurrentCount = PSH_Globals.FormCurrentCount - 1;
            PSH_Globals.FormTypeList.Remove((PSH_Globals.FormTypeListCount - 1).ToString());
            PSH_Globals.FormTypeListCount = PSH_Globals.FormTypeListCount - 1;

            oTempClass = null;
        }

        /// <summary>
        /// 폼현재객체수
        /// </summary>
        /// <returns></returns>
        public static int Get_CurrentFormsCount()
        {
            return PSH_Globals.FormCurrentCount;
        }

        /// <summary>
        /// 폼총객체수
        /// </summary>
        /// <returns></returns>
        public static int Get_TotalFormsCount()
        {
            return PSH_Globals.FormTotalCount;
        }
    }
}

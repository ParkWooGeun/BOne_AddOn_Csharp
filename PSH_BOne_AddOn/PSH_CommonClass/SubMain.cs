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

            while (MessageAPIs.GetMessage(ref MessageAPIs.structMsg, 0, 0, 0))
            {
                MessageAPIs.TranslateMessage(ref MessageAPIs.structMsg);
                MessageAPIs.DispatchMessage(ref MessageAPIs.structMsg);
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
        public static void Add_Forms(object cObject, string oFormUid, object oFormTypeEx)
        {
            PSH_Globals.ClassList.Add(cObject, oFormUid);
            PSH_Globals.FormTotalCount += 1;
            PSH_Globals.FormCurrentCount += 1;
            PSH_Globals.FormTypeListCount += 1;
            PSH_Globals.FormTypeList.Add(oFormTypeEx, oFormUid);
        }

        /// <summary>
        /// 폼객체 제거
        /// </summary>
        /// <param name="oFormUniqueID"></param>
        public static void Remove_Forms(string oFormUniqueID)
        {
            try
            {
                PSH_Globals.ClassList.Remove(oFormUniqueID);
                PSH_Globals.FormTotalCount -= 1;
                PSH_Globals.FormCurrentCount -= 1;
                PSH_Globals.FormTypeListCount -= 1;
                PSH_Globals.FormTypeList.Remove(oFormUniqueID);
            }
            catch(System.Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(ex.Message);
            }
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

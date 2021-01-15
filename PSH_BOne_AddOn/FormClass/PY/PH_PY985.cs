using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 의료비지급명세서자료 전산매체수록
    /// </summary>
    internal class PH_PY985 : PSH_BaseClass
    {
        public string oFormUniqueID01;
        public override void LoadForm()
        {
            string strXml = string.Empty;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY985.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY985_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY985");

                strXml = oXmlDoc.xml.ToString();
                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                CreateItems();

                oForm.EnableMenu("1281", false); //찾기
                oForm.EnableMenu("1282", true); //추가
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", false); //행삭제
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void CreateItems()
        {   
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                oForm.Items.Item("YYYY").Specific.Value = DateTime.Now.AddYears(-1).ToString("yyyy"); //년도 기본년도에서 - 1

                oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 10); //제출일자
                oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 신고파일 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create()
        {
            bool functionReturnValue = false;
            
            short errNum = 0;
            string stringSpace = string.Empty;
            string cltcod = string.Empty;
            string yyyy = string.Empty;
            string hTaxID = string.Empty;
            string docDate = string.Empty;

            try
            {
                cltcod = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("YYYY").Specific.Value.ToString().Trim();
                hTaxID = oForm.Items.Item("HtaxID").Specific.Value.ToString().Trim();
                docDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

                if (PSH_Globals.SBO_Application.MessageBox("의료비 신고파일을 생성하시겠습니까?", 2, "&Yes!", "&No") == 2)
                {
                    errNum = 1;
                    throw new Exception();
                }
                
                if (File_Create_A_record(cltcod, yyyy, hTaxID, docDate) == false)  //A RECORD 처리
                {
                    errNum = 2;
                    throw new Exception();
                }

                FileSystem.FileClose(1);

                PSH_Globals.SBO_Application.StatusBar.SetText("전산매체수록이 정상적으로 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                functionReturnValue = true;
            }
            catch(Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("취소하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("A레코드 생성 실패.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    stringSpace = new string(' ', 10);
                    PSH_Globals.SBO_Application.StatusBar.SetText("File_Create 실행 중 오류가 발생했습니다." + stringSpace + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
            
            return functionReturnValue;
        }

        /// <summary>
        /// A레코드 생성
        /// </summary>
        /// <returns></returns>
        private bool File_Create_A_record(string pcltcod, string pStdYear, string phTaxID, string pdocDate)
        {
            bool functionReturnValue = false;
            short errNum = 0;
            string sQry = string.Empty;
            string saup = string.Empty;
            int newCNT = 0; //일련번호
            string oFilePath = string.Empty; //파일 경로

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //2013년기준 250 BYTE
            //2015년기준 251 BYTE
            //2017년기준 251 BYTE
            //2019년기준 236 BYTE
            //2020년기준 236 BYTE

            string A001; // 레코드구분(1) 'A'
            string A002; // 자료구분(2)   '26'
            string A003; // 세무서(3)
            string A004; // 일련번호(6)
            string A005; // 제출년월일(8)
            string A006; // 사업자번호(10)
            string A007; // 홈텍스ID(20)
            string A008; // 세무프로그램코드(4)
            string A009; // 귀속년도(4)
            string A010; // 사업자번호(10)
            string A011; // 법인명(상호)(40)
            string A012; // 소득자의주민등록번호(13)
            string A013; // 내,외국인(1)
            string A014; // 성명(30)
            string A015; // 지급처사업자등록번호(10)
            string A016; // 지급처상호(40)
            string A017; // 의료증빙코드(1)
            string A018; // 건수(5)
            string A019; // 지급금액(11)
            string A020; // 난임시술비해당여부(1)
            string A021; // 주민등록번호(13)
            string A022; // 내,외국인코드(1)
            string A023; // 본인등해당여부(1)
            string A024; // 제출대상기간코드(1)

            try
            {   
                //A_RECORE QUERY
                sQry = "      EXEC PH_PY985_A '";
                sQry += pcltcod + "', '";
                sQry += pStdYear + "', '";
                sQry += phTaxID + "', '";
                sQry += pdocDate + "'";
                oRecordSet.DoQuery(sQry);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("작성시작!", oRecordSet.RecordCount, false);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }
                else
                {
                    //PATH및 파일이름 만들기
                    saup = oRecordSet.Fields.Item("A010").Value; //사업자번호
                    oFilePath = "C:\\BANK\\CA" + codeHelpClass.Mid(saup, 0, 7) + "." + codeHelpClass.Mid(saup, 7, 3);
                    FileSystem.FileClose(1);
                    FileSystem.FileOpen(1, oFilePath, OpenMode.Output);

                    while (!oRecordSet.EoF)
                    {
                        newCNT += 1; //일련번호

                        //A RECORD MOVE
                        A001 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A001").Value.ToString().Trim(), 1);
                        A002 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A002").Value.ToString().Trim(), 2);
                        A003 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A003").Value.ToString().Trim(), 3);
                        A004 = codeHelpClass.GetFixedLengthStringByte(newCNT.ToString(), 6, '0');
                        A005 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A005").Value.ToString().Trim(), 8);
                        A006 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A006").Value.ToString().Trim(), 10);
                        A007 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A007").Value.ToString().Trim(), 20);
                        A008 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A008").Value.ToString().Trim(), 4);
                        A009 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A009").Value.ToString().Trim(), 4);
                        A010 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A010").Value.ToString().Trim(), 10);
                        A011 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A011").Value.ToString().Trim(), 40);
                        A012 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A012").Value.ToString().Trim(), 13);
                        A013 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A013").Value.ToString().Trim(), 1);
                        A014 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A014").Value.ToString().Trim(), 30);
                        A015 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A015").Value.ToString().Trim(), 10);
                        A016 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A016").Value.ToString().Trim(), 40);
                        A017 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A017").Value.ToString().Trim(), 1);
                        A018 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A018").Value.ToString().Trim(), 5, '0');
                        A019 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A019").Value.ToString().Trim(), 11, '0');
                        A020 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A020").Value.ToString().Trim(), 1);
                        A021 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A021").Value.ToString().Trim(), 13);
                        A022 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A022").Value.ToString().Trim(), 1);
                        A023 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A023").Value.ToString().Trim(), 1);
                        A024 = codeHelpClass.GetFixedLengthStringByte(oRecordSet.Fields.Item("A024").Value.ToString().Trim(), 1);

                        FileSystem.PrintLine(1, A001 + A002 + A003 + A004 + A005 + A006 + A007 + A008 + A009 + A010 + A011 + A012 + A013 + A014 + A015 + A016 + A017 + A018 + A019 + A020 + A021 + A022 + A023 + A024);

                        oRecordSet.MoveNext();

                        ProgressBar01.Value = ProgressBar01.Value + 1;
                        ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 작성중........!";
                    }
                }

                functionReturnValue = true;
            }
            catch(Exception ex)
            {
                ProgressBar01.Stop();

                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("의료비자료가 존재하지 않습니다. 등록하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 필수 입력값 체크
        /// </summary>
        /// <returns></returns>
        private bool HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            short errNum = 0;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("HtaxID").Specific.VALUE))
                {
                    errNum = 1;
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.VALUE))
                {
                    errNum = 2;
                    throw new Exception();
                }

                functionReturnValue = true;
            }
            catch(Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("홈텍스ID(5자리이상)를 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("제출일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return functionReturnValue;
        }

        

        /// <summary>
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (File_Create() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        else
                        {
                            BubbleEvent = false;
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
        
        /// <summary>
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            //사업장 변경되면
                            case "CLTCOD":
                                sQry =  "        SELECT      U_HomeTId,";
                                sQry += "             U_ChgDpt,";
                                sQry += "             U_ChgName,";
                                sQry += "             U_ChgTel";
                                sQry += " FROM        [@PH_PY005A]";
                                sQry += " WHERE       U_CLTCode = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("HtaxID").Specific.Value = oRecordSet.Fields.Item("U_HomeTId").Value.ToString().Trim();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// FORM_UNLOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Form Item Event
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">pVal</param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    break;
            }
        }

        /// <summary>
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1287": //복제
                            break;
                        case "1281":
                        case "1282":
                            break;
                        case "1288": // TODO: to "1291"
                            break;
                        case "1293":
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:   //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:    //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:   //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:    //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}

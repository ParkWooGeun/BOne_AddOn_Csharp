using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 교대근무인정현황
    /// </summary>
    internal class PH_PY683 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        //'// 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.DataTable oDS_PH_PY683;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY683.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY683_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY683");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY683_CreateItems();
                PH_PY683_FormItemEnabled();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY683_CreateItems()
        {
            try
            {
                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY683");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY683");
                oDS_PH_PY683 = oForm.DataSources.DataTables.Item("PH_PY683");

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 기준년도
                oForm.DataSources.UserDataSources.Add("YM", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 10);
                oForm.Items.Item("YM").Specific.DataBind.SetBound(true, "", "YM");
                oForm.DataSources.UserDataSources.Item("YM").Value = DateTime.Now.ToString("yyyyMM");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY683_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        public void PH_PY683_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY683_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_FormItemEvent
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">이벤트 </param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                    ////2
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                    ////3
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                    ////4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    // Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                    ////7
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                    ////8
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                    ////9
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
                    ////12
                    break;


                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                    ////16
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                    ////18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                    ////19
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                    ////20
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    // Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
                    ////22
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
                    ////23
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    // Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
                    ////37
                    break;

                case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
                    ////38
                    break;

                case SAPbouiCOM.BoEventTypes.et_Drag:
                    ////39
                    break;

            }
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
                    if (pVal.ItemUID == "BtnSearch")
                    {
                        if (PH_PY683_DataValidCheck() == true)
                        {
                            PH_PY683_DataFind();
                        }
                        else
                        {
                            BubbleEvent = false;
                        }
                    }

                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY683_DataValidCheck()
        {
            bool functionReturnValue = false;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.VALUE))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("년월운 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY683_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                functionReturnValue = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        /// <returns></returns>
        private void PH_PY683_DataFind()
        {
            int iRow = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string YM = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();
                YM = oForm.Items.Item("YM").Specific.VALUE.Trim();

                sQry = "Exec PH_PY683_01 '" + CLTCOD + "','" + YM + "'";

                oDS_PH_PY683.ExecuteQuery(sQry);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY683_DataFind_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

    }
}



//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
// // ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_HR_Addon
//{
//	internal class PH_PY683
//	{
////****************************************************************************************************************
//////  File : PH_PY683.cls
//////  Module : 인사관리 > 근태관리 > 근태리포트
//////  Desc : 교대근무인정현황
//////  FormType : PH_PY683
//////  Create Date(Start) : 2016.01.28
//////  Creator : NGY
//////  Modified Date :
//////  Modifier :
//////  Company : Poongsan Holdings
////****************************************************************************************************************

//		public string oFormUniqueID01;
//		public SAPbouiCOM.Form oForm;
//		public SAPbouiCOM.Grid oGrid01;

//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string oLastItemUID01;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//		private string oLastColUID01;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//		private int oLastColRow01;

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm(string oFromDocEntry01 = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			string oInnerXml = null;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY683.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

//			//매트릭스의 타이틀높이와 셀높이를 고정
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}

//			oFormUniqueID01 = "PH_PY683_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID01, "PH_PY683");
//			////폼추가
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID01);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			////oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

//			oForm.Freeze(true);
//			PH_PY683_CreateItems();
//			PH_PY683_ComboBox_Setting();
//			PH_PY683_CF_ChooseFromList();
//			PH_PY683_EnableMenus();
//			PH_PY683_SetDocument(oFromDocEntry01);
//			PH_PY683_FormResize();

//			oForm.EnableMenu("1283", false);
//			//삭제
//			oForm.EnableMenu("1286", false);
//			//닫기
//			oForm.EnableMenu("1287", false);
//			//복제
//			oForm.EnableMenu("1285", false);
//			//복원
//			oForm.EnableMenu("1284", false);
//			//취소
//			oForm.EnableMenu("1293", false);
//			//행삭제
//			oForm.EnableMenu("1281", false);
//			oForm.EnableMenu("1282", true);



//			oForm.Update();
//			oForm.Freeze(false);

//			oForm.Visible = true;
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;

//			return;
//			LoadForm_Error:
//			oForm.Update();
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oForm = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY683_MTX01()
//		{
//			//******************************************************************************
//			//Function ID : PH_PY683_MTX01()
//			//해당모듈 : PH_PY683
//			//기능 : 데이터 조회
//			//인수 : 없음
//			//반환값 : 없음
//			//특이사항 : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short i = 0;
//			string sQry = null;
//			short ErrNum = 0;

//			//    Dim RecordSet01 As SAPbobsCOM.Recordset
//			//    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)

//			string CLTCOD = null;
//			string YM = null;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			YM = Strings.Trim(oForm.Items.Item("YM").Specific.VALUE);

//			oForm.Freeze(true);

//			sQry = "EXEC [PH_PY683_01] '" + CLTCOD + "', '" + YM + "'";

//			oGrid01.DataTable = oForm.DataSources.DataTables.Item("DataTable");
//			oGrid01.DataTable.Clear();
//			oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry);

//			if (oGrid01.Rows.Count == 0) {
//				ErrNum = 1;
//				goto PH_PY683_MTX01_Error;
//			}

//			oGrid01.AutoResizeColumns();
//			oForm.Update();

//			oForm.Freeze(false);

//			//    Set ProgBar01 = Nothing
//			//    Set RecordSet01 = Nothing
//			return;
//			PH_PY683_MTX01_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//    ProgBar01.Stop
//			oForm.Freeze(false);
//			//    Set ProgBar01 = Nothing
//			//    Set RecordSet01 = Nothing

//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.", ref "W");
//			} else {
//				MDC_Com.MDC_GF_Message(ref "PH_PY683_MTX01_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//		}

//		private bool PH_PY683_HeaderSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			//******************************************************************************
//			//Function ID : PH_PY683_HeaderSpaceLineDel()
//			//해당모듈 : PH_PY683
//			//기능 : 필수입력사항 체크
//			//인수 : 없음
//			//반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
//			//특이사항 : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short ErrNum = 0;
//			ErrNum = 0;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			switch (true) {
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("YM").Specific.VALUE)):
//					ErrNum = 1;
//					goto PH_PY683_HeaderSpaceLineDel_Error;
//					break;

//			}

//			functionReturnValue = true;
//			return functionReturnValue;
//			PH_PY683_HeaderSpaceLineDel_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "기준년월는 필수조회 조건입니다. 확인하세요.", ref "E");
//				oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

///// 메트릭스 필수 사항 check
//		private bool PH_PY683_MatrixSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			int i = 0;
//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			string sQry = null;

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			functionReturnValue = true;
//			return functionReturnValue;
//			PH_PY683_MatrixSpaceLineDel_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 2) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 사원코드가 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 3) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 시간이 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 4) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 등록일자가 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 5) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 비가동코드가 없습니다. 확인하세요.", ref "E");
//			} else {
//				MDC_Com.MDC_GF_Message(ref "PH_PY683_MatrixSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private void PH_PY683_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short i = 0;
//			short ErrNum = 0;
//			string sQry = null;
//			string ItemCode = null;

//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string CLTCOD = null;
//			string TeamCode = null;
//			string RspCode = null;

//			SAPbouiCOM.ProgressBar ProgBar01 = null;

//			oForm.Freeze(true);

//			switch (oUID) {



//			}

//			oForm.Freeze(false);
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			return;
//			PH_PY683_FlushToItemValue_Error:

//			oForm.Freeze(false);
//			//    Call ProgBar01.Stop
//			//    ProgBar01.VALUE = 100
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;

//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "PH_PY683_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}

//		}

/////폼의 아이템 사용지정
//		public void PH_PY683_FormItemEnabled()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//				//        Call CLTCOD_Select(oForm, "SCLTCOD")

//				//        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
//				//        oForm.Items("ItemCode").Enabled = True

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//				//        Call CLTCOD_Select(oForm, "SCLTCOD")

//				//        oForm.Items("ItemCode").Enabled = True

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//				//        Call CLTCOD_Select(oForm, "SCLTCOD")

//			}

//			return;
//			PH_PY683_FormItemEnabled_Error:

//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			MDC_Com.MDC_GF_Message(ref "PH_PY683_FormItemEnabled_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

/////아이템 변경 이벤트
//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1
//					Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					////5
//					Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
//					Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//					////8
//					Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					////10
//					Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//					////18
//					break;
//				////et_FORM_ACTIVATE
//				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//					////19
//					break;
//				////et_FORM_DEACTIVATE
//				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//					////20
//					Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//					////27
//					Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//					////4
//					break;
//				////et_LOST_FOCUS
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					////17
//					Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//			}
//			return;
//			Raise_FormItemEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			////BeforeAction = True
//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
//					case "1284":
//						//취소
//						break;
//					case "1286":
//						//닫기
//						break;
//					case "1293":
//						//행삭제
//						break;
//					case "1281":
//						//찾기
//						break;
//					case "1282":
//						//추가
//						///추가버튼 클릭시 메트릭스 insertrow

//						//                Call PH_PY683_FormReset

//						//                oMat01.Clear
//						//                oMat01.FlushToDataSource
//						//                oMat01.LoadFromDataSource

//						//                oForm.Mode = fm_ADD_MODE
//						//                BubbleEvent = False
//						//                Call PH_PY683_LoadCaption

//						//oForm.Items("GCode").Click ct_Regular


//						return;

//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						//레코드이동버튼
//						break;

//					case "7169":
//						//엑셀 내보내기
//						break;

//				}
//			////BeforeAction = False
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1284":
//						//취소
//						break;
//					case "1286":
//						//닫기
//						break;
//					case "1293":
//						//행삭제
//						break;
//					case "1281":
//						//찾기
//						break;
//					////Call PH_PY683_FormItemEnabled '//UDO방식
//					case "1282":
//						//추가
//						break;
//					//                oMat01.Clear
//					//                oDS_PH_PY683A.Clear

//					//                Call PH_PY683_LoadCaption
//					//                Call PH_PY683_FormItemEnabled
//					////Call PH_PY683_FormItemEnabled '//UDO방식
//					////Call PH_PY683_AddMatrixRow(0, True) '//UDO방식
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						//레코드이동버튼
//						break;
//					////Call PH_PY683_FormItemEnabled

//					case "7169":
//						//엑셀 내보내기
//						break;

//				}
//			}
//			return;
//			Raise_FormMenuEvent_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			////BeforeAction = True
//			if ((BusinessObjectInfo.BeforeAction == true)) {
//				switch (BusinessObjectInfo.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//						////33
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//						////34
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//						////35
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//						////36
//						break;
//				}
//			////BeforeAction = False
//			} else if ((BusinessObjectInfo.BeforeAction == false)) {
//				switch (BusinessObjectInfo.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//						////33
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//						////34
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//						////35
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//						////36
//						break;
//				}
//			}
//			return;
//			Raise_FormDataEvent_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//			}
//			if (pval.ItemUID == "Mat01") {
//				if (pval.Row > 0) {
//					oLastItemUID01 = pval.ItemUID;
//					oLastColUID01 = pval.ColUID;
//					oLastColRow01 = pval.Row;
//				}
//			} else {
//				oLastItemUID01 = pval.ItemUID;
//				oLastColUID01 = "";
//				oLastColRow01 = 0;
//			}
//			return;
//			Raise_RightClickEvent_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//				if (pval.ItemUID == "PH_PY683") {
//					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//					}
//				}

//				///조회
//				if (pval.ItemUID == "BtnSearch") {

//					if (PH_PY683_HeaderSpaceLineDel() == false) {
//						BubbleEvent = false;
//						return;
//					}

//					PH_PY683_MTX01();

//					//        ElseIf pval.ItemUID = "BtnPrint" Then
//					//
//					//            If PH_PY683_HeaderSpaceLineDel = False Then
//					//                BubbleEvent = False
//					//                Exit Sub
//					//            End If
//					//
//					//            Call PH_PY683_Print_Report01

//				}

//			} else if (pval.BeforeAction == false) {
//				if (pval.ItemUID == "PH_PY683") {
//					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//					}
//				}
//			}

//			return;
//			Raise_EVENT_ITEM_PRESSED_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//				MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "MSTCOD", "");
//				//사번

//			} else if (pval.BeforeAction == false) {

//			}

//			return;
//			Raise_EVENT_KEY_DOWN_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {

//			}

//			return;
//			Raise_EVENT_CLICK_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {

//				PH_PY683_FlushToItemValue(pval.ItemUID);

//			}

//			return;
//			Raise_EVENT_COMBO_SELECT_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {

//			}
//			return;
//			Raise_EVENT_DOUBLE_CLICK_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {

//			}
//			return;
//			Raise_EVENT_MATRIX_LINK_PRESSED_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			if (pval.BeforeAction == true) {

//				if (pval.ItemChanged == true) {

//					PH_PY683_FlushToItemValue(pval.ItemUID);

//				}

//			} else if (pval.BeforeAction == false) {

//			}

//			oForm.Freeze(false);

//			return;
//			Raise_EVENT_VALIDATE_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {
//				PH_PY683_FormItemEnabled();
//				////Call PH_PY683_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
//			}
//			return;
//			Raise_EVENT_MATRIX_LOAD_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {
//				PH_PY683_FormResize();
//			}
//			return;
//			Raise_EVENT_RESIZE_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {
//				//        If (pval.ItemUID = "ItemCode") Then
//				//            Dim oDataTable01 As SAPbouiCOM.DataTable
//				//            Set oDataTable01 = pval.SelectedObjects
//				//            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
//				//            Set oDataTable01 = Nothing
//				//        End If
//				//        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
//				//            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY683A", "U_CardCode,U_CardName")
//				//        End If
//			}
//			return;
//			Raise_EVENT_CHOOSE_FROM_LIST_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.ItemUID == "Mat01") {
//				if (pval.Row > 0) {
//					oLastItemUID01 = pval.ItemUID;
//					oLastColUID01 = pval.ColUID;
//					oLastColRow01 = pval.Row;
//				}
//			} else {
//				oLastItemUID01 = pval.ItemUID;
//				oLastColUID01 = "";
//				oLastColRow01 = 0;
//			}
//			return;
//			Raise_EVENT_GOT_FOCUS_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//				SubMain.RemoveForms(oFormUniqueID01);
//				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oForm = null;
//				//UPGRADE_NOTE: oGrid01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oGrid01 = null;
//			}
//			return;
//			Raise_EVENT_FORM_UNLOAD_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			if ((oLastColRow01 > 0)) {
//				if (pval.BeforeAction == true) {

//				} else if (pval.BeforeAction == false) {

//				}
//			}
//			return;
//			Raise_EVENT_ROW_DELETE_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private bool PH_PY683_CreateItems()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			string oQuery01 = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oGrid01 = oForm.Items.Item("Grid01").Specific;
//			oForm.DataSources.DataTables.Add("DataTable");
//			oGrid01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;

//			//사업장
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");

//			//기준년월
//			oForm.DataSources.UserDataSources.Add("YM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("YM").Specific.DataBind.SetBound(true, "", "YM");

//			//UPGRADE_WARNING: oForm.Items(YM).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("YM").Specific.String = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM");

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY683_CreateItems_Error:

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY683_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

/////콤보박스 set
//		public void PH_PY683_ComboBox_Setting()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbouiCOM.ComboBox oCombo = null;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;

//			oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);

//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;

//			return;
//			PH_PY683_ComboBox_Setting_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY683_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY683_CF_ChooseFromList()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			return;
//			PH_PY683_CF_ChooseFromList_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY683_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY683_EnableMenus()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			return;
//			PH_PY683_EnableMenus_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY683_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY683_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY683_FormItemEnabled();
//				////Call PH_PY683_AddMatrixRow(0, True) '//UDO방식일때
//			} else {
//				//        oForm.Mode = fm_FIND_MODE
//				//        Call PH_PY683_FormItemEnabled
//				//        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
//				//        oForm.Items("1").Click ct_Regular
//			}
//			return;
//			PH_PY683_SetDocument_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY683_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY683_FormResize()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			return;
//			PH_PY683_FormResize_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY683_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

////Private Sub PH_PY683_Print_Report01()
////    On Error GoTo PH_PY683_Print_Report01_Error
////
////    Dim WinTitle As String
////    Dim ReportName As String
////    Dim sQry As String
////
////    Dim CLTCOD As String '사업장
////    Dim TeamCode As String '팀
////    Dim RspCode As String '담당
////    Dim ClsCode As String '반
////    Dim MSTCOD As String '사원번호
////    Dim yyyy As String '기준년도
////
////    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
////    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE) '팀
////    RspCode = Trim(oForm.Items("RspCode").Specific.VALUE) '담당
////    ClsCode = Trim(oForm.Items("ClsCode").Specific.VALUE) '반
////    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE) '사원번호
////    yyyy = Trim(oForm.Items("YYYY").Specific.VALUE) '기준년도
////
////    Dim ProgBar01 As SAPbouiCOM.ProgressBar
////    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
////
////     '/ ODBC 연결 체크
////    If ConnectODBC = False Then
////        GoTo PH_PY683_Print_Report01_Error
////    End If
////
////    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
////
////    WinTitle = "[PH_PY683] 비근무일수 현황"
////
////    ReportName = "PH_PY683_01.rpt"
////
////    '// Formula 수식필드
////    ReDim gRpt_Formula(1)
////    ReDim gRpt_Formula_Value(1)
////
////    '// SubReport
////    ReDim gRpt_SRptSqry(1)
////    ReDim gRpt_SRptName(1)
////
////    ReDim gRpt_SFormula(1, 1)
////    ReDim gRpt_SFormula_Value(1, 1)
////
////    gRpt_SFormula(1, 1) = ""
////    gRpt_SFormula_Value(1, 1) = ""
////
////    '/ Procedure 실행"
////    sQry = "                EXEC [PH_PY683_02] "
////    sQry = sQry & "'" & CLTCOD & "'," '사업장
////    sQry = sQry & "'" & TeamCode & "',"  '팀
////    sQry = sQry & "'" & RspCode & "',"  '담당
////    sQry = sQry & "'" & ClsCode & "',"  '반
////    sQry = sQry & "'" & MSTCOD & "',"  '사번
////    sQry = sQry & "'" & yyyy & "'"  '기준년도
////
////    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V", "", 2) = False Then
////        GoTo PH_PY683_Print_Report01_Error
////    End If
////
////    ProgBar01.VALUE = 100
////    ProgBar01.Stop
////    Set ProgBar01 = Nothing
////
////    Exit Sub
////
////PH_PY683_Print_Report01_Error:
////
////    ProgBar01.VALUE = 100
////    ProgBar01.Stop
////    Set ProgBar01 = Nothing
////
////    Call MDC_Com.MDC_GF_Message("Print_Query_Error:" & Err.Number & " - " & Err.Description, "E")
////
////End Sub
//	}
//}

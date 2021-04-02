using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
namespace MDC_PS_Addon
{
	internal class PS_PP480
	{
//****************************************************************************************************************
////  File           : PS_PP480.cls
////  Module         : PP
////  Description    : 생산완료 현황
////  FormType       : PS_PP480
////  Create Date    : 2011.02.17
////  Modified Date  :
////  Creator        : Lee Byong Gak
////  Company        : Poongsan Holdings
//****************************************************************************************************************

		public string oFormUniqueID01;
		public SAPbouiCOM.Form oForm01;

//****************************************************************************************************************
// .srf 파일로부터 폼을 로드한다.
//****************************************************************************************************************
		public void LoadForm()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			string oInnerXml01 = null;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			oXmlDoc01.load(SubMain.ShareFolderPath + "ScreenPS\\PS_PP480.srf");
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

			//매트릭스의 타이틀높이와 셀높이를 고정
			for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
			}

			oFormUniqueID01 = "PS_PP480_" + GetTotalFormsCount();
			SubMain.AddForms(this, oFormUniqueID01);
			////폼추가
			SubMain.Sbo_Application.LoadBatchActions(out (oXmlDoc01.xml));

			//폼 할당
			oForm01 = SubMain.Sbo_Application.Forms.Item(oFormUniqueID01);

			oForm01.SupportedModes = -1;
			oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

			//////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//************************************************************************************************************
			//화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
			//    oForm01.DataBrowser.BrowseBy = "DocNum"
			//************************************************************************************************************
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////

			oForm01.Freeze(true);
			CreateItems();
			ComboBox_Setting();
			Initialization();

			oForm01.EnableMenu(("1283"), false);
			//// 삭제
			oForm01.EnableMenu(("1286"), false);
			//// 닫기
			oForm01.EnableMenu(("1287"), false);
			//// 복제
			oForm01.EnableMenu(("1284"), true);
			//// 취소
			oForm01.EnableMenu(("1293"), false);
			//// 행삭제

			oForm01.Update();
			oForm01.Freeze(false);
			oForm01.Visible = true;

			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oXmlDoc01 = null;
			return;
			LoadForm_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			oForm01.Update();
			oForm01.Freeze(false);
			//UPGRADE_NOTE: oXmlDoc01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oXmlDoc01 = null;
			if ((oForm01 == null) == false) {
				//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				oForm01 = null;
			}
			MDC_Com.MDC_GF_Message(ref "LoadForm_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void Initialization()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			SAPbouiCOM.ComboBox oCombo = null;

			////아이디별 사업장 세팅
			oCombo = oForm01.Items.Item("BPLId").Specific;
			oCombo.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);

			////아이디별 사번 세팅
			//    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD

			////아이디별 부서 세팅
			//    Set oCombo = oForm01.Items("DeptCode").Specific
			//    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue

			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("SjAmt").Specific.VALUE = 0;

			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oCombo = null;
			return;
			Initialization_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oCombo = null;
			MDC_Com.MDC_GF_Message(ref "Initialization_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		private void CreateItems()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////디비데이터 소스 개체 할당
			//    Set oDS_PS_PP480H = oForm01.DataSources.DBDataSources("@PS_PP480H")
			//    Set oDS_PS_PP480L = oForm01.DataSources.DBDataSources("@PS_PP480L")

			//// 메트릭스 개체 할당
			//    Set oMat01 = oForm01.Items("Mat01").Specific

			//사업장
			oForm01.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

			//기간Fr
			oForm01.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 10);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
			oForm01.DataSources.UserDataSources.Item("DocDateFr").Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");

			//기간To
			oForm01.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 10);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
			oForm01.DataSources.UserDataSources.Item("DocDateTo").Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");

			//품목코드
			oForm01.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

			//대분류
			oForm01.DataSources.UserDataSources.Add("OrdGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("OrdGbn").Specific.DataBind.SetBound(true, "", "OrdGbn");

			//거래처코드
			oForm01.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

			//거래처명
			oForm01.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

			//자체/외주구분
			oForm01.DataSources.UserDataSources.Add("InOutGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("InOutGbn").Specific.DataBind.SetBound(true, "", "InOutGbn");

			//수주금액
			oForm01.DataSources.UserDataSources.Add("SjAmt", SAPbouiCOM.BoDataType.dt_SUM);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("SjAmt").Specific.DataBind.SetBound(true, "", "SjAmt");

			//연간품 제외(작번등록 기준)
			oForm01.DataSources.UserDataSources.Add("YearPd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("YearPd").Specific.DataBind.SetBound(true, "", "YearPd");
			//UPGRADE_WARNING: oForm01.Items().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("YearPd").Specific.Checked = true;

			return;
			CreateItems_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "CreateItems_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void ComboBox_Setting()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////콤보에 기본값설정
			SAPbouiCOM.ComboBox oCombo = null;
			string sQry = null;
			SAPbobsCOM.Recordset oRecordSet01 = null;

			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			//// 사업장
			oCombo = oForm01.Items.Item("BPLId").Specific;
			sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
			oRecordSet01.DoQuery(sQry);
			while (!(oRecordSet01.EoF)) {
				oCombo.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
				oRecordSet01.MoveNext();
			}

			//자체/외주구분
			oCombo = oForm01.Items.Item("InOutGbn").Specific;
			oCombo.ValidValues.Add("%", "전체");
			oCombo.ValidValues.Add("IN", "자체");
			oCombo.ValidValues.Add("OUT", "외주");
			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

			//    Call oForm01.Freeze(True)

			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oCombo = null;
			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet01 = null;
			return;
			ComboBox_Setting_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oCombo = null;
			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet01 = null;
			MDC_Com.MDC_GF_Message(ref "ComboBox_Setting_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

//****************************************************************************************************************
//// ItemEventHander
//****************************************************************************************************************
		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			int ErrNum = 0;
			SAPbouiCOM.ProgressBar ProgressBar01 = null;

			////BeforeAction = True
			if ((pval.BeforeAction == true)) {
				switch (pval.EventType) {
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
						////1
						if (pval.ItemUID == "1") {
							if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
								//                        If HeaderSpaceLineDel = False Then
								//                            BubbleEvent = False
								//                            Exit Sub
								//                        End If
								//                        If MatrixSpaceLineDel = False Then
								//                            BubbleEvent = False
								//                            Exit Sub
								//                        End If
							}

						//출력버튼 클릭시
						} else if (pval.ItemUID == "Btn01") {
							if (HeaderSpaceLineDel() == false) {
								BubbleEvent = false;
								return;
							} else {
								Print_Query();
							}
						}
						break;
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
						////2
						if (pval.CharPressed == 9) {
						}
						break;
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
						////5
						break;
					case SAPbouiCOM.BoEventTypes.et_CLICK:
						////6
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
						////7
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
						////8
						break;
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
						////10                                  '질의 창 띄워서 명칭 넣어주기
						break;
					//                    If pval.ItemUID = "ItmMsort" Then
					//                       FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
					//                    End If
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
						////11
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
						////18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
						////19
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
						////20
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
						////27
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
						////3
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
						////4
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
						////17
						break;
				}

				//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			////BeforeAction = False
			} else if ((pval.BeforeAction == false)) {
				switch (pval.EventType) {
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
						////1
						break;
					//                If pval.ItemUID = "1" Then
					//                    If oForm01.Mode = fm_ADD_MODE Then
					//                        oForm01.Mode = fm_OK_MODE
					//                        Call Sbo_Application.ActivateMenuItem("1282")
					//                    ElseIf oForm01.Mode = fm_OK_MODE Then
					//                        FormItemEnabled
					//                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
					//                    End If
					//                End If
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
						////2
						break;
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
						////5
						break;
					case SAPbouiCOM.BoEventTypes.et_CLICK:
						////6
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
						////7
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
						////8
						break;
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
						////10
						if (pval.ItemUID == "CardCode") {
							FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);
						}
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
						////11
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
						////18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
						////19
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
						////20
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
						////27
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
						////3
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
						////4
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
						////17
						SubMain.RemoveForms(oFormUniqueID01);
						//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						oForm01 = null;
						break;
					//                Set oMat01 = Nothing
				}
			}
			return;
			Raise_ItemEvent_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			ProgressBar01 = null;
			if (ErrNum == 101) {
				ErrNum = 0;
				MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
				BubbleEvent = false;
			} else {
				MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
			}
		}

		public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;

			////BeforeAction = True
			if ((pval.BeforeAction == true)) {
				switch (pval.MenuUID) {
					case "1284":
						//취소
						break;
					case "1286":
						//닫기
						break;
					case "1293":
						//행삭제
						break;
					case "1281":
						//찾기
						break;
					case "1282":
						//추가
						break;
					case "1285":
						//복원
						break;
					case "1288":
					case "1289":
					case "1290":
					case "1291":
						//레코드이동버튼
						break;
				}

				//-----------------------------------------------------------------------------------------------------------
			////BeforeAction = False
			} else if ((pval.BeforeAction == false)) {
				switch (pval.MenuUID) {
					case "1284":
						//취소
						break;
					case "1286":
						//닫기
						break;
					case "1285":
						//복원
						break;
					case "1293":
						//행삭제
						break;
					case "1281":
						//찾기
						break;
					case "1282":
						//추가
						break;
					case "1288":
					case "1289":
					case "1290":
					case "1291":
						//레코드이동버튼
						break;
				}
			}
			return;
			Raise_MenuEvent_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			if ((eventInfo.BeforeAction == true)) {

			} else if ((eventInfo.BeforeAction == false)) {
				////작업
			}
			return;
			Raise_RightClickEvent_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		}

		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////BeforeAction = True
			if ((BusinessObjectInfo.BeforeAction == true)) {
				switch (BusinessObjectInfo.EventType) {
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
						////33
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
						////34
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
						////35
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
						////36
						break;
				}
			////BeforeAction = False
			} else if ((BusinessObjectInfo.BeforeAction == false)) {
				switch (BusinessObjectInfo.EventType) {
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
						////33
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
						////34
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
						////35
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
						////36
						break;
				}
			}
			return;
			Raise_FormDataEvent_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		private void FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			string DocNum = null;
			string LineId = null;
			short ErrNum = 0;
			string sQry = null;
			SAPbobsCOM.Recordset oRecordSet = null;

			//UPGRADE_NOTE: TypeName이(가) TypeName_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			string ItemType = null;
			string Mark = null;
			string ItemMsort = null;
			string CardItmMsort = null;
			string CardName = null;
			string MsortName = null;
			string MarkName = null;
			string TypeName_Renamed = null;

			oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			//--------------------------------------------------------------
			//Header--------------------------------------------------------
			switch (oUID) {
				case "OrdGbn":
					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sQry = "SELECT Name FROM [@PSH_ITMBSORT] WHERE Code =  '" + Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.VALUE) + "'";
					oRecordSet.DoQuery(sQry);

					//UPGRADE_WARNING: oForm01.Items(CodeName).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					oForm01.Items.Item("CodeName").Specific.String = Strings.Trim(oRecordSet.Fields.Item("Name").Value);
					break;
				case "CardCode":
					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sQry = "SELECT CardName FROM [OCRD] WHERE CardCode =  '" + Strings.Trim(oForm01.Items.Item("CardCode").Specific.VALUE) + "'";
					oRecordSet.DoQuery(sQry);

					//UPGRADE_WARNING: oForm01.Items(CardName).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					oForm01.Items.Item("CardName").Specific.String = Strings.Trim(oRecordSet.Fields.Item("CardName").Value);
					break;
			}


			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet = null;
			return;
			FlushToItemValue_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			MDC_Com.MDC_GF_Message(ref "FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
		}

		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			 // ERROR: Not supported in C#: OnErrorStatement

			short ErrNum = 0;

			ErrNum = 0;

			//// Check
			switch (true) {
				//        Case Trim(oDS_PS_PP480H.GetValue("U_BPLId", 0)) = ""
				//            ErrNum = 1
				//            GoTo HeaderSpaceLineDel_Error
			}

			functionReturnValue = true;
			return functionReturnValue;
			HeaderSpaceLineDel_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			if (ErrNum == 1) {
				MDC_Com.MDC_GF_Message(ref "사업장은 필수사항입니다. 확인하여 주십시오.", ref "E");
			} else {
				MDC_Com.MDC_GF_Message(ref "HeaderSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
			}
			functionReturnValue = false;
			return functionReturnValue;
		}

		private void Print_Query()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			short i = 0;
			short ErrNum = 0;
			string WinTitle = null;
			string ReportName = null;
			string sQry = null;
			string Sub_sQry = null;

			string BPLID = null;
			string DocDateFr = null;
			string DocDateTo = null;
			string ItemCode = null;
			string OrdGbn = null;
			string CardCode = null;

			string InOutGbn = null;
			decimal SjAmt = default(decimal);
			string YearPd = null;

			SAPbobsCOM.Recordset oRecordSet = null;

			oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			SAPbouiCOM.ProgressBar ProgBar01 = null;
			ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

			MDC_PS_Common.ConnectODBC();

			//// 조회조건문

			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BPLID = Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DocDateFr = Strings.Trim(oForm01.Items.Item("DocDateFr").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DocDateTo = Strings.Trim(oForm01.Items.Item("DocDateTo").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ItemCode = Strings.Trim(oForm01.Items.Item("ItemCode").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OrdGbn = Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CardCode = Strings.Trim(oForm01.Items.Item("CardCode").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			InOutGbn = Strings.Trim(oForm01.Items.Item("InOutGbn").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SjAmt = oForm01.Items.Item("SjAmt").Specific.VALUE;

			//연간품제외 여부
			if (oForm01.DataSources.UserDataSources.Item("YearPd").Value == "Y") {
				YearPd = "N";
			} else {
				YearPd = "Y";
			}

			if (string.IsNullOrEmpty(ItemCode))
				ItemCode = "%";
			if (string.IsNullOrEmpty(OrdGbn))
				OrdGbn = "%";
			if (string.IsNullOrEmpty(CardCode))
				CardCode = "%";

			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

			WinTitle = "[PS_PP480_01] 생산완료 현황";
			ReportName = "PS_PP480_01.RPT";
			MDC_Globals.gRpt_Formula = new string[3];
			MDC_Globals.gRpt_Formula_Value = new string[3];

			//// Formula 수식필드

			MDC_Globals.gRpt_Formula[1] = "DocDateFr";
			MDC_Globals.gRpt_Formula_Value[1] = Strings.Left(DocDateFr, 4) + "-" + Strings.Mid(DocDateFr, 5, 2) + "-" + Strings.Right(DocDateFr, 2);
			MDC_Globals.gRpt_Formula[2] = "DocDateTo";
			MDC_Globals.gRpt_Formula_Value[2] = Strings.Left(DocDateTo, 4) + "-" + Strings.Mid(DocDateTo, 5, 2) + "-" + Strings.Right(DocDateTo, 2);
			MDC_Globals.gRpt_SRptSqry = new string[2];
			MDC_Globals.gRpt_SRptName = new string[2];
			MDC_Globals.gRpt_SFormula = new string[2, 2];
			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

			//// SubReport


			MDC_Globals.gRpt_SFormula[1, 1] = "";
			MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

			/// Procedure 실행"

			sQry = "         EXEC [PS_PP480_01] '";
			sQry = sQry + BPLID + "', '";
			sQry = sQry + DocDateFr + "', '";
			sQry = sQry + DocDateTo + "', '";
			sQry = sQry + ItemCode + "', '";
			sQry = sQry + OrdGbn + "', '";
			sQry = sQry + CardCode + "','";
			sQry = sQry + InOutGbn + "','";
			sQry = sQry + SjAmt + "','";
			sQry = sQry + YearPd + "'";

			//    oRecordSet.DoQuery sQry
			//    If oRecordSet.RecordCount = 0 Then
			//        ErrNum = 1
			//        GoTo Print_Query_Error
			//    End If

			/// Action (sub_query가 있을때는 'Y'로...)/
			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") == false) {
			}

			ProgBar01.Value = 100;
			ProgBar01.Stop();
			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			ProgBar01 = null;

			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet = null;
			return;
			Print_Query_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

			ProgBar01.Value = 100;
			ProgBar01.Stop();
			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			ProgBar01 = null;

			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet = null;

			if (ErrNum == 1) {
				MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다. 확인해 주세요.", ref "E");
			} else {
				MDC_Com.MDC_GF_Message(ref "Print_Query_Error:" + Err().Number + " - " + Err().Description, ref "E");
			}
		}
	}
}

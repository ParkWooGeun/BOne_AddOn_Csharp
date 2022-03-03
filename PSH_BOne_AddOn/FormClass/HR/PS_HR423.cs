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
	internal class PS_HR423
	{
//****************************************************************************************************************
////  File           : PS_HR423.cls
////  Module         : HR
////  Description    : 전문직정량평가현황
////  FormType       : PS_HR423
////  Create Date    : 2013.06.19
////  Modified Date  :
////  Creator        : NGY
////  Company        : Poongsan Holdings
//****************************************************************************************************************

		public string oFormUniqueID01;
		public SAPbouiCOM.Form oForm01;
		public SAPbouiCOM.Matrix oMat01;
			//등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_HR423H;
			//등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_HR423L;

			//클래스에서 선택한 마지막 아이템 Uid값
		private string oLast_Item_UID;
			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private string oLast_Col_UID;
			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_Col_Row;

		private int oLast_Mode;

//****************************************************************************************************************
// .srf 파일로부터 폼을 로드한다.
//****************************************************************************************************************
		public void LoadForm()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			int i = 0;
			string oInnerXml01 = null;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			oXmlDoc01.load(SubMain.ShareFolderPath + "ScreenPS\\PS_HR423.srf");
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
			oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

			//매트릭스의 타이틀높이와 셀높이를 고정
			for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
				oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
			}

			oFormUniqueID01 = "PS_HR423_" + GetTotalFormsCount();
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

			oForm01.EnableMenu(("1283"), false);
			//// 삭제
			oForm01.EnableMenu(("1286"), false);
			//// 닫기
			oForm01.EnableMenu(("1287"), false);
			//// 복제
			oForm01.EnableMenu(("1284"), false);
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

		private void CreateItems()
		{
			 // ERROR: Not supported in C#: OnErrorStatement

			////디비데이터 소스 개체 할당
			//    Set oDS_PS_HR423H = oForm01.DataSources.DBDataSources("@PS_HR423H")
			//    Set oDS_PS_HR423L = oForm01.DataSources.DBDataSources("@PS_HR423L")

			//// 메트릭스 개체 할당
			//    Set oMat01 = oForm01.Items("Mat01").Specific

			oForm01.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
			//UPGRADE_WARNING: oForm01.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oForm01.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");
			oForm01.DataSources.UserDataSources.Item("Year").Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY");

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
			sQry = "SELECT BPLId, BPLName From OBPL Order by BPLId";
			oRecordSet01.DoQuery(sQry);
			while (!(oRecordSet01.EoF)) {
				oCombo.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
				oRecordSet01.MoveNext();
			}
			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

			//// 평가차수
			oCombo = oForm01.Items.Item("Number").Specific;
			oCombo.ValidValues.Add("1", "1차평가");
			oCombo.ValidValues.Add("2", "2차평가");
			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

			//// 평가그룹
			oCombo = oForm01.Items.Item("Group").Specific;
			oCombo.ValidValues.Add("1", "반장");
			oCombo.ValidValues.Add("2", "사원");
			oCombo.ValidValues.Add("3", "임금피크");
			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

			//oForm01.Items("BPLId").Click ct_Regular

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
							////헤더
							//                    If pval.ItemUID = "ItmBsort" Then
							//                        If oForm01.Items("ItmBsort").Specific.VALUE = "" Then
							//                            Sbo_Application.ActivateMenuItem ("7425")
							//                            BubbleEvent = False
							//                        End If
							//                    End If
							if (pval.ItemUID == "ItmMsort") {
								//UPGRADE_WARNING: oForm01.Items(ItmMsort).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								if (string.IsNullOrEmpty(oForm01.Items.Item("ItmMsort").Specific.VALUE)) {
									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
									BubbleEvent = false;
								}
							}
							////라인
							//                    If pval.ItemUID = "Mat01" Then
							//                        If pval.ColUID = "PP070No" Then
							//                            If oMat01.Columns("PP070No").Cells(pval.Row).Specific.Value = "" Then
							//                                Sbo_Application.ActivateMenuItem ("7425")
							//                                BubbleEvent = False
							//                            End If
							//                        End If
							//                    End If
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
						////10
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

			oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			//--------------------------------------------------------------
			//Header--------------------------------------------------------
			switch (oUID) {

			}

			//--------------------------------------------------------------
			//Line----------------------------------------------------------
			if (oUID == "Mat01") {
				switch (oCol) {

				}
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
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			switch (true) {
				case string.IsNullOrEmpty(Strings.Trim(oForm01.Items.Item("Year").Specific.VALUE)):
					ErrNum = 1;
					goto HeaderSpaceLineDel_Error;
					break;
				case string.IsNullOrEmpty(Strings.Trim(oForm01.Items.Item("Number").Specific.VALUE)):
					ErrNum = 2;
					goto HeaderSpaceLineDel_Error;
					break;
				case Strings.Len(Strings.Trim(oForm01.Items.Item("Group").Specific.VALUE)) == Convert.ToDouble(""):
					ErrNum = 3;
					goto HeaderSpaceLineDel_Error;
					break;

			}

			functionReturnValue = true;
			return functionReturnValue;
			HeaderSpaceLineDel_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			if (ErrNum == 1) {
				MDC_Com.MDC_GF_Message(ref "년도는 필수사항입니다. 확인하여 주십시오.", ref "E");
			} else if (ErrNum == 2) {
				MDC_Com.MDC_GF_Message(ref "평가차수는 필수사항입니다. 확인하여 주십시오.", ref "E");
			} else if (ErrNum == 3) {
				MDC_Com.MDC_GF_Message(ref "평가그룹을 확인하여 주십시오.", ref "E");

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

			//UPGRADE_NOTE: Year이(가) Year_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			string Year_Renamed = null;
			string Number = null;
			string Group = null;
			string BPLID = null;

			SAPbobsCOM.Recordset oRecordSet = null;

			oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			MDC_PS_Common.ConnectODBC();

			//// 조회조건문
			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BPLID = Strings.Trim(oForm01.Items.Item("BPLId").Specific.Selected.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Year_Renamed = Strings.Trim(oForm01.Items.Item("Year").Specific.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Number = Strings.Trim(oForm01.Items.Item("Number").Specific.Selected.VALUE);
			//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Group = Strings.Trim(oForm01.Items.Item("Group").Specific.Selected.VALUE);

			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
			WinTitle = "[PS_HR423] 전문직정량평가현황";
			ReportName = "PS_HR423_01.RPT";
			MDC_Globals.gRpt_Formula = new string[5];
			MDC_Globals.gRpt_Formula_Value = new string[5];

			//// Formula 수식필드

			MDC_Globals.gRpt_Formula[1] = "BPLId";
			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" + Strings.Trim(oForm01.Items.Item("BPLId").Specific.VALUE) + "'";
			oRecordSet.DoQuery(sQry);
			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MDC_Globals.gRpt_Formula_Value[1] = oRecordSet.Fields.Item(0).Value;

			MDC_Globals.gRpt_Formula[2] = "Year";
			MDC_Globals.gRpt_Formula_Value[2] = Year_Renamed;
			MDC_Globals.gRpt_Formula[3] = "Number";
			MDC_Globals.gRpt_Formula_Value[3] = Number;
			MDC_Globals.gRpt_Formula[4] = "Group";
			if (Group == "1") {
				MDC_Globals.gRpt_Formula_Value[4] = "반장";
			} else {
				MDC_Globals.gRpt_Formula_Value[4] = "사원";
			}
			MDC_Globals.gRpt_SRptSqry = new string[2];
			MDC_Globals.gRpt_SRptName = new string[2];
			MDC_Globals.gRpt_SFormula = new string[2, 2];
			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];


			//// SubReport


			MDC_Globals.gRpt_SFormula[1, 1] = "";
			MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

			/// Procedure 실행"
			sQry = "EXEC [PS_HR423_01] '" + BPLID + "','" + Year_Renamed + "','" + Number + "','" + Group + "'";
			oRecordSet.DoQuery(sQry);
			if (oRecordSet.RecordCount == 0) {
				ErrNum = 1;
				goto Print_Query_Error;
			}

			/// Action (sub_query가 있을때는 'Y'로...)/
			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") == false) {
			}

			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oRecordSet = null;
			return;
			Print_Query_Error:
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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

using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 여신한도 초과요청
	/// </summary>
	internal class PS_SD080 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01; 
		private SAPbouiCOM.DBDataSource oDS_PS_SD080H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD080L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_Mode;
		private int oSeq;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
		public override void LoadForm(string oFormDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD080.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD080_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD080");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocNum";

				oForm.Freeze(true);

				PS_SD080_CreateItems();
				PS_SD080_SetComboBox();
                PS_SD080_Initialize();
                PS_SD080_ClearForm();
                //PS_SD080_EnableFormItem();
                //oDS_PS_SD080H.SetValue("U_OKYN", 0, "N");

                oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1286", true); //닫기
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1284", true); //취소
				oForm.EnableMenu("1293", true); //행삭제
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
			}
		}

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_SD080_CreateItems()
        {
            try
            {
                //디비데이터 소스 개체 할당
                oDS_PS_SD080H = oForm.DataSources.DBDataSources.Item("@PS_SD080H");
                oDS_PS_SD080L = oForm.DataSources.DBDataSources.Item("@PS_SD080L");

                //메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;

                oDS_PS_SD080H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 콤보박스 세팅
        /// </summary>
        private void PS_SD080_SetComboBox()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 화면 초기화
        /// </summary>
        private void PS_SD080_Initialize()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DocNum 초기화
        /// </summary>
        private void PS_SD080_ClearForm()
        {
            string DocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD080'", "");
                if (string.IsNullOrEmpty(DocNum) || DocNum == "0")
                {
                    oForm.Items.Item("DocNum").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocNum").Specific.Value = DocNum;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	int ErrNum = 0;
        //	object TempForm01 = null;
        //	SAPbouiCOM.ProgressBar ProgressBar01 = null;

        //	string ItemType = null;
        //	string RequestDate = null;
        //	string Size = null;
        //	string ItemCode = null;
        //	string ItemName = null;
        //	string Unit = null;
        //	string DueDate = null;
        //	string RequestNo = null;
        //	int Qty = 0;
        //	decimal Weight = default(decimal);
        //	string RFC_Sender = null;
        //	double Calculate_Weight = 0;
        //	int Seq = 0;

        //	////BeforeAction = True
        //	if ((pVal.BeforeAction == true)) {
        //		switch (pVal.EventType) {
        //			//et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1
        //				if (pVal.ItemUID == "1") {
        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //						if (PS_SD080_DeleteHeaderSpaceLine() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}
        //						if (PS_SD080_DeleteMatrixSpaceLine() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}
        //						//                        Call PS_SD080_DeleteEmptyRow
        //						oLast_Mode = oForm.Mode;
        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
        //						oLast_Mode = oForm.Mode;
        //					}
        //				}
        //				break;
        //			//et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //				////2
        //				if (pVal.CharPressed == 9) {
        //					if (pVal.ItemUID == "CntcCode") {
        //						//UPGRADE_WARNING: oForm.Items(CntcCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value)) {
        //							SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //							BubbleEvent = false;
        //						}
        //					}
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //				////5
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CLICK:
        //				////6
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //				////7
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //				////8
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //				////10
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //				////11
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //				////18
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //				////19
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //				////20
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //				////27
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //				////3
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //				////4
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //				////17
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((pVal.BeforeAction == false)) {
        //		switch (pVal.EventType) {
        //			//et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1
        //				if (pVal.ItemUID == "1") {
        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == true) {
        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //						SubMain.Sbo_Application.ActivateMenuItem("1282");
        //						//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("BPLId").Specific.Select("4", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //					} else if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
        //						PS_SD080_EnableFormItem();
        //						oLast_Mode = 100;
        //					}
        //				} else if (pVal.ItemUID == "Btn01") {
        //					//                    If oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value = "" Then
        //					TempForm01 = new PS_SD081();
        //					//UPGRADE_WARNING: TempForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					TempForm01.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
        //					BubbleEvent = false;
        //					//                    End If
        //				} else if (pVal.ItemUID == "Btn02") {
        //					PS_SD080_PrintReport01();
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //				////2
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //				////5
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CLICK:
        //				////6
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //				////7
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //				////8
        //				break;
        //			//et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //				////10
        //				if (pVal.ItemChanged == true) {
        //					if (pVal.ItemUID == "CntcCode") {
        //						PS_SD080_FlushToItemValue(pVal.ItemUID);
        //					}
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //				////11
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //				////18
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //				////19
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //				////20
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //				////27
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //				////3
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //				////4
        //				break;
        //			//et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //				////17
        //				SubMain.RemoveForms(oFormUniqueID);
        //				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oForm = null;
        //				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oMat01 = null;
        //				//UPGRADE_NOTE: oDS_PS_SD080H 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oDS_PS_SD080H = null;
        //				//UPGRADE_NOTE: oDS_PS_SD080L 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oDS_PS_SD080L = null;
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_ItemEvent_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	if (ErrNum == 101) {
        //		ErrNum = 0;
        //		MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //		BubbleEvent = false;
        //	} else {
        //		MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;

        //	////BeforeAction = True
        //	if ((pVal.BeforeAction == true)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				break;
        //			case "1281":
        //				//찾기
        //				break;
        //			case "1282":
        //				//추가
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((pVal.BeforeAction == false)) {
        //		switch (pVal.MenuUID) {
        //			//[1284:취소] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case "1284":
        //				//취소
        //				PS_SD080_EnableFormItem();
        //				oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			//[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case "1293":
        //				//행삭제
        //				if (oMat01.RowCount != oMat01.VisualRowCount) {
        //					for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
        //					}

        //					oMat01.FlushToDataSource();
        //					oDS_PS_SD080L.RemoveRecord(oDS_PS_SD080L.Size - 1);
        //					//// Mat01에 마지막라인(빈라인) 삭제
        //					oMat01.Clear();
        //					oMat01.LoadFromDataSource();
        //				}
        //				break;
        //			//[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case "1281":
        //				//찾기
        //				PS_SD080_EnableFormItem();
        //				oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				break;
        //			//[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case "1282":
        //				//추가
        //				PS_SD080_EnableFormItem();
        //				PS_SD080_ClearForm();
        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("BPLId").Specific.Select("4", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //				oDS_PS_SD080H.SetValue("U_DocDate", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD"));
        //				oDS_PS_SD080H.SetValue("U_OKYN", 0, "N");
        //				break;
        //			//                PS_SD080_AddMatrixRow 0, True
        //			//                oForm.Items("BPLId").Click ct_Collapsed
        //			//[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				PS_SD080_EnableFormItem();
        //				break;
        //			//                If oMat01.VisualRowCount > 0 Then
        //			//                    If oMat01.Columns("CGNo").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
        //			//                        If oDS_PS_SD080H.GetValue("Status", 0) = "O" Then
        //			//                            PS_SD080_AddMatrixRow oMat01.RowCount, False
        //			//                        End If
        //			//                    End If
        //			//                End If
        //		}
        //	}
        //	return;
        //	Raise_MenuEvent_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((BusinessObjectInfo.BeforeAction == true)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_FormDataEvent_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if ((eventInfo.BeforeAction == true)) {
        //		////작업
        //	} else if ((eventInfo.BeforeAction == false)) {
        //		////작업
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion





        #region CF_ChooseFromList
        //public void CF_ChooseFromList()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////ChooseFromList 설정
        //	return;
        //	CF_ChooseFromList_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	MDC_Com.MDC_GF_Message(ref "CF_ChooseFromList_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD080_EnableFormItem
        //public void PS_SD080_EnableFormItem()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //		oForm.Items.Item("DocNum").Enabled = false;
        //		oForm.Items.Item("Btn01").Enabled = true;
        //		oForm.Items.Item("BPLId").Enabled = true;
        //		oForm.Items.Item("CntcCode").Enabled = true;
        //		oForm.Items.Item("DocDate").Enabled = true;
        //		oMat01.Columns.Item("RequestP").Editable = true;
        //		oMat01.Columns.Item("Comment").Editable = true;
        //	} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
        //		oForm.Items.Item("DocNum").Enabled = true;
        //		oForm.Items.Item("Btn01").Enabled = false;
        //		oForm.Items.Item("BPLId").Enabled = true;
        //		oForm.Items.Item("CntcCode").Enabled = true;
        //		oForm.Items.Item("DocDate").Enabled = true;
        //		oMat01.Columns.Item("RequestP").Editable = true;
        //		oMat01.Columns.Item("Comment").Editable = true;
        //	} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //		oForm.Items.Item("DocNum").Enabled = false;
        //		oForm.Items.Item("Btn01").Enabled = false;
        //		if (Strings.Trim(oDS_PS_SD080H.GetValue("U_OkYN", 0)) == "Y") {
        //			oForm.Items.Item("BPLId").Enabled = false;
        //			oForm.Items.Item("CntcCode").Enabled = false;
        //			oForm.Items.Item("DocDate").Enabled = false;
        //			oMat01.Columns.Item("RequestP").Editable = false;
        //			oMat01.Columns.Item("Comment").Editable = false;
        //		} else {
        //			oForm.Items.Item("BPLId").Enabled = true;
        //			oForm.Items.Item("CntcCode").Enabled = true;
        //			oForm.Items.Item("DocDate").Enabled = true;
        //			oMat01.Columns.Item("RequestP").Editable = true;
        //			oMat01.Columns.Item("Comment").Editable = true;
        //		}
        //	}
        //	return;
        //	PS_SD080_EnableFormItem_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	MDC_Com.MDC_GF_Message(ref "PS_SD080_EnableFormItem_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion


        #region PS_SD080_AddMatrixRow
        //public void PS_SD080_AddMatrixRow(int oRow, ref bool RowIserted = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////행추가여부
        //	if (RowIserted == false) {
        //		oDS_PS_SD080L.InsertRecord((oRow));
        //	}
        //	oMat01.AddRow();
        //	oDS_PS_SD080L.Offset = oRow;
        //	oDS_PS_SD080L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
        //	oMat01.LoadFromDataSource();
        //	return;
        //	PS_SD080_AddMatrixRow_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	MDC_Com.MDC_GF_Message(ref "PS_SD080_AddMatrixRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD080_FlushToItemValue
        //private void PS_SD080_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	short ErrNum = 0;
        //	string sQry = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;

        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	switch (oUID) {
        //		case "CntcCode":
        //			sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + Strings.Trim(oDS_PS_SD080H.GetValue("U_CntcCode", 0)) + "'";
        //			oRecordSet01.DoQuery(sQry);

        //			oDS_PS_SD080H.SetValue("U_CntcName", 0, Strings.Trim(oRecordSet01.Fields.Item(0).Value));
        //			break;
        //	}

        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	return;
        //	PS_SD080_FlushToItemValue_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	MDC_Com.MDC_GF_Message(ref "PS_SD080_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD080_DeleteHeaderSpaceLine
        //private bool PS_SD080_DeleteHeaderSpaceLine()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	short ErrNum = 0;

        //	ErrNum = 0;

        //	//// Check
        //	switch (true) {
        //		case string.IsNullOrEmpty(oDS_PS_SD080H.GetValue("U_BPLId", 0)):
        //			ErrNum = 1;
        //			goto PS_SD080_DeleteHeaderSpaceLine_Error;
        //			break;
        //		case string.IsNullOrEmpty(oDS_PS_SD080H.GetValue("U_CntcCode", 0)):
        //			ErrNum = 2;
        //			goto PS_SD080_DeleteHeaderSpaceLine_Error;
        //			break;
        //		case string.IsNullOrEmpty(oDS_PS_SD080H.GetValue("U_DocDate", 0)):
        //			ErrNum = 3;
        //			goto PS_SD080_DeleteHeaderSpaceLine_Error;
        //			break;
        //	}

        //	functionReturnValue = true;
        //	return functionReturnValue;
        //	PS_SD080_DeleteHeaderSpaceLine_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	if (ErrNum == 1) {
        //		MDC_Com.MDC_GF_Message(ref "사업장은 필수사항입니다. 확인하세요.", ref "E");
        //	} else if (ErrNum == 2) {
        //		MDC_Com.MDC_GF_Message(ref "작성자는 필수사항입니다. 확인하세요.", ref "E");
        //	} else if (ErrNum == 3) {
        //		MDC_Com.MDC_GF_Message(ref "요청일자는 필수사항입니다. 확인하세요.", ref "E");
        //	} else {
        //		MDC_Com.MDC_GF_Message(ref "PS_SD080_DeleteHeaderSpaceLine_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //	functionReturnValue = false;
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD080_DeleteMatrixSpaceLine
        //private bool PS_SD080_DeleteMatrixSpaceLine()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	short ErrNum = 0;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	string sQry = null;

        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	ErrNum = 0;

        //	oMat01.FlushToDataSource();

        //	//// 라인
        //	if (oMat01.VisualRowCount == 0) {
        //		ErrNum = 1;
        //		goto PS_SD080_DeleteMatrixSpaceLine_Error;
        //	}

        //	for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //		if (Convert.ToDouble(oDS_PS_SD080L.GetValue("U_RequestP", i)) <= 0) {
        //			ErrNum = 3;
        //			goto PS_SD080_DeleteMatrixSpaceLine_Error;
        //		}
        //	}

        //	oMat01.LoadFromDataSource();

        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	functionReturnValue = true;
        //	return functionReturnValue;
        //	PS_SD080_DeleteMatrixSpaceLine_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	if (ErrNum == 1) {
        //		MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하세요.", ref "E");
        //	} else if (ErrNum == 2) {
        //		MDC_Com.MDC_GF_Message(ref "초과필요금액은 0 보다 커야 합니다. 확인하세요.", ref "E");
        //	} else {
        //		MDC_Com.MDC_GF_Message(ref "PS_SD080_DeleteMatrixSpaceLine_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //	functionReturnValue = false;
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD080_DeleteEmptyRow
        //public void PS_SD080_DeleteEmptyRow()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;

        //	oMat01.FlushToDataSource();

        //	for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //		if (string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD080L.GetValue("U_CGNo", i)))) {
        //			oDS_PS_SD080L.RemoveRecord(i);
        //			//// Mat01에 마지막라인(빈라인) 삭제
        //		}
        //	}

        //	oMat01.LoadFromDataSource();
        //	return;
        //	PS_SD080_DeleteEmptyRow_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	MDC_Com.MDC_GF_Message(ref "PS_SD080_DeleteEmptyRow_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD080_PrintReport01
        //private void PS_SD080_PrintReport01()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	short i = 0;
        //	short ErrNum = 0;
        //	string WinTitle = null;
        //	string ReportName = null;

        //	string DocDateTo = null;
        //	string CntcCode = null;
        //	string BPLId = null;
        //	string DocDateFr = null;
        //	string PackNo = null;
        //	string DocNum = null;

        //	string sQry = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;

        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	MDC_PS_Common.ConnectODBC();

        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocNum = Strings.Trim(oForm.Items.Item("DocNum").Specific.Value);

        //	WinTitle = "[PS_SD080]" + "여신한도 초과 승인 신청서";
        //	ReportName = "PS_SD080_01.RPT";
        //	MDC_Globals.gRpt_Formula = new string[2];
        //	MDC_Globals.gRpt_Formula_Value = new string[2];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];

        //	//// Formula 수식필드

        //	//// SubReport

        //	//// 조회조건문
        //	oMat01.FlushToDataSource();

        //	sQry = "EXEC [PS_SD080_02] '" + DocNum + "'";
        //	oRecordSet01.DoQuery(sQry);
        //	if (oRecordSet01.RecordCount == 0) {
        //		ErrNum = 1;
        //		goto PS_SD080_PrintReport01_Error;
        //	}

        //	//// Action
        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") == false) {
        //	}

        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	return;
        //	PS_SD080_PrintReport01_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	if (ErrNum == 1) {
        //		MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다.확인해 주세요.", ref "E");
        //	} else {
        //		MDC_Com.MDC_GF_Message(ref "PS_SD080_PrintReport01_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //}
        #endregion
    }
}

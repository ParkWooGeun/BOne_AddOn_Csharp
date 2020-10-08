using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 활동유형등록
    /// </summary>
    internal class PS_CO050 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;

        private SAPbouiCOM.DBDataSource oDS_PS_CO050H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_CO050L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO050.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_CO050_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_CO050");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                CreateItems();
                AddMatrixRow(0, true);
                ComboBox_Setting();

                oForm.EnableMenu("1293", true); // 행삭제
                oForm.EnableMenu("1287", false); // 복제
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void CreateItems()
        {
            try
            {
                //oForm.Freeze(true);

                oDS_PS_CO050H = oForm.DataSources.DBDataSources.Item("@PS_CO050H");
                oDS_PS_CO050L = oForm.DataSources.DBDataSources.Item("@PS_CO050L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                //oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 행추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false) //행추가여부
                {
                    oRow = oMat01.RowCount;
                    oDS_PS_CO050L.InsertRecord((oRow));
                }

                oMat01.AddRow();
                oDS_PS_CO050L.Offset = oRow;
                oDS_PS_CO050L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 콤보박스 세팅
        /// </summary>
        private void ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //단위
                sQry = "SELECT U_Minor,U_CdName From [@PS_SY001L] Where Code = 'C001' Order by U_Minor";
                oRecordSet01.DoQuery(sQry);

                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("Unit").Cells.Item(1).Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                oMat01.Columns.Item("Unit").Cells.Item(1).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //2차원가요소
                sQry = "SELECT U_CECode, U_CEName From [@PS_CO010L] Where U_Category= '99'";
                oRecordSet01.DoQuery(sQry);

                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("SecondCE").Cells.Item(1).Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                oMat01.Columns.Item("SecondCE").Cells.Item(1).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 필수입력사항 체크(Header)
        /// </summary>
        /// <returns></returns>
        private bool HeaderSpaceLineDel()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            
            try
            {
                if (string.IsNullOrEmpty(oDS_PS_CO050H.GetValue("Code", 0)) || string.IsNullOrEmpty(oDS_PS_CO050H.GetValue("Name", 0)))
                {
                    errCode = "1";
                    throw new Exception();
                }
            
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("원가영역 및 또는 원가영역명은 필수입력 사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            
            return returnValue;
        }


        /// <summary>
        /// 필수입력사항 체크(Line)
        /// </summary>
        /// <returns></returns>
        private bool MatrixSpaceLineDel()
        {
            bool returnValue = false;
            int i;
            string errCode = string.Empty;

            try
            {
                oMat01.FlushToDataSource();

                if (oMat01.VisualRowCount < 1) //라인
                {
                    errCode = "1";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount > 0)
                {
                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                    {
                        oDS_PS_CO050L.Offset = i;

                        if (string.IsNullOrEmpty(oDS_PS_CO050L.GetValue("U_ActCode", i)))
                        {
                            errCode = "2";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_CO050L.GetValue("U_ActName", i)))
                        {
                            errCode = "3";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_CO050L.GetValue("U_Unit", i)))
                        {
                            errCode = "4";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_CO050L.GetValue("U_SecondCE", i)))
                        {
                            errCode = "5";
                            throw new Exception();
                        }
                    }

                    if (string.IsNullOrEmpty(oDS_PS_CO050L.GetValue("U_ActCode", oMat01.VisualRowCount - 1)))
                    {
                        oDS_PS_CO050L.RemoveRecord(oMat01.VisualRowCount - 1);
                    }
                }

                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("라인데이타가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("활동유형코드는 필수입력사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("활동유형이름은 필수입력사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("단위는 필수입력사항입니다. 확인하세요", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "5")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("2차원가요소는 필수입력사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            
            return returnValue;
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void FlushToItemValue(string oUID, int oRow, string oCol)
        {
            try
            {
                if (oUID == "oMat01")
                {
                    oMat01.FlushToDataSource();

                    if (oCol == "ActCode")
                    {
                        oDS_PS_CO050L.Offset = oRow - 1;
                            
                        if (oRow == oMat01.RowCount && !string.IsNullOrEmpty(oDS_PS_CO050L.GetValue("U_ActCode", oRow - 1).ToString().Trim()))
                        {
                            AddMatrixRow(0, false);
                            oMat01.Columns.Item("ActCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                    }

                    oMat01.LoadFromDataSource();
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        





        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((pval.BeforeAction == true)) {
        //		switch (pval.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1
        //				if (pval.ItemUID == "1") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //						if (HeaderSpaceLineDel() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}

        //						if (MatrixSpaceLineDel() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}
        //						//                      ElseIf oForm01.Mode = fm_FIND_MODE Then
        //					}
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
        //	} else if ((pval.BeforeAction == false)) {
        //		switch (pval.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1
        //				if (pval.ItemUID == "1") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //						oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //						SubMain.Sbo_Application.ActivateMenuItem("1282");
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //						FormItemEnabled();
        //						AddMatrixRow(0, ref true);
        //					}
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //				////2
        //				break;
        //			//                 Call ComboBox_Setting
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
        //				//활동유형코드가 바뀌면 한 줄이 더 생기도록
        //				if (pval.ItemChanged == true) {
        //					if (pval.ColUID == "ActCode") {
        //						FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);
        //					}
        //				}
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //				////11
        //				AddMatrixRow(pval.Row, ref false);
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
        //				SubMain.RemoveForms(oFormUniqueID01);
        //				//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oForm01 = null;
        //				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oMat01 = null;
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_ItemEvent_Error:
        //	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	int i = 0;

        //	////BeforeAction = True
        //	if ((pval.BeforeAction == true)) {
        //		switch (pval.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
        //				break;
        //			case "1281":
        //				//찾기
        //				oForm01.DataBrowser.BrowseBy = "Code";
        //				////UDO방식일때
        //				break;
        //			case "1282":
        //				//추가
        //				oForm01.DataBrowser.BrowseBy = "Code";
        //				////UDO방식일때
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((pval.BeforeAction == false)) {
        //		switch (pval.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				if (oMat01.RowCount != oMat01.VisualRowCount) {
        //					for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
        //					}
        //					oMat01.FlushToDataSource();
        //					// DBDataSource에 레코드가 한줄 더 생긴다.
        //					oDS_PS_CO050L.RemoveRecord(oDS_PS_CO050L.Size - 1);
        //					// 레코드 한 줄을 지운다.
        //					oMat01.LoadFromDataSource();
        //					// DBDataSource를 매트릭스에 올리고
        //					if (oMat01.RowCount == 0) {
        //						//                        Call AddMatrixRow(1, 0, True)
        //						AddMatrixRow(1, ref 0);
        //					} else {
        //						if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_CO050L.GetValue("U_ActCode", oMat01.RowCount - 1)))) {
        //							//                            Call AddMatrixRow(1, oMat01.RowCount, True)
        //							AddMatrixRow(1, ref oMat01.RowCount);
        //						}
        //					}
        //				}
        //				break;
        //			case "1281":
        //				//찾기
        //				AddMatrixRow(0, ref true);
        //				////UDO방식
        //				oForm01.DataBrowser.BrowseBy = "Code";
        //				////UDO방식일때        '찾기버튼 클릭시 Matrix에 행 추가
        //				break;
        //			case "1282":
        //				//추가
        //				//                 Call FormItemEnabled '//UDO방식
        //				AddMatrixRow(0, ref true);
        //				////UDO방식
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼             '추가버튼 클릭시 Matrix에 행 추가
        //				break;
        //			////Call FormItemEnabled
        //		}
        //	}
        //	return;
        //	MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	if ((oLastColRow01 > 0)) {
        //		if (pval.BeforeAction == true) {
        //			////행삭제전 행삭제가능여부검사
        //		} else if (pval.BeforeAction == false) {
        //			for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
        //			}
        //			oMat01.FlushToDataSource();
        //			oDS_PS_CO050L.RemoveRecord(oDS_PS_CO050L.Size - 1);
        //			oMat01.LoadFromDataSource();
        //			if (oMat01.RowCount == 0) {
        //				AddMatrixRow(0);
        //			} else {
        //				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_CO050L.GetValue("U_CECode", oMat01.RowCount - 1)))) {
        //					AddMatrixRow(oMat01.RowCount);
        //				}
        //			}
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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


    }
}

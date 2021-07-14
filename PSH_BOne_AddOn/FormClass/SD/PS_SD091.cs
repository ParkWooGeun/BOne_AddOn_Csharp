using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 이동요청등록
	/// </summary>
	internal class PS_SD091 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_SD091H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD091L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD091.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD091_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD091");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

				PS_SD091_CreateItems();
				PS_SD091_SetComboBox();
				PS_SD091_Initialize();
				PS_SD091_CF_ChooseFromList();
				PS_SD091_EnableMenus();
				PS_SD091_SetDocument(oFormDocEntry);
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
        private void PS_SD091_CreateItems()
        {
            try
            {
                oDS_PS_SD091H = oForm.DataSources.DBDataSources.Item("@PS_SD091H");
                oDS_PS_SD091L = oForm.DataSources.DBDataSources.Item("@PS_SD091L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");

                oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SD091_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL Where BPLId = '1' Or BPLId = '4' order by BPLId", "1", false, false);
                dataHelpClass.Set_ComboList(oForm.Items.Item("OutWhCd").Specific, "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", "104", false, true);
                dataHelpClass.Set_ComboList(oForm.Items.Item("InWhCd").Specific, "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", "101", false, true);
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] order by Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmMsort"), "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] order by U_Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] order by Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] order by Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Mark"), "SELECT Code, Name FROM [@PSH_MARK] order by Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("SbasUnit"), "SELECT Code, Name  FROM [@PSH_UOMORG] order by Code", "", "");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면 초기화
        /// </summary>
        private void PS_SD091_Initialize()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                string lcl_User_BPLId = dataHelpClass.User_BPLID();
                
                if (lcl_User_BPLId == "1" || lcl_User_BPLId == "4") //소속사업장이 창원이나 구로일 때만 콤보박스 Select
                {
                    oForm.Items.Item("BPLId").Specific.Select(lcl_User_BPLId, SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD(); // 인수자
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList 설정
        /// </summary>
        private void PS_SD091_CF_ChooseFromList()
        {
            SAPbouiCOM.ChooseFromList oCFL01 = null;
            SAPbouiCOM.ChooseFromList oCFL02 = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;
            SAPbouiCOM.ChooseFromListCollection oCFLs01 = null;
            SAPbouiCOM.ChooseFromListCollection oCFLs02 = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams01 = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams02 = null;
            SAPbouiCOM.EditText oEdit01 = null;
            SAPbouiCOM.EditText oEdit02 = null;

            try
            {
                oEdit01 = oForm.Items.Item("CardCode").Specific;
                oCFLs01 = oForm.ChooseFromLists;
                oCFLCreationParams01 = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams01.ObjectType = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                oCFLCreationParams01.UniqueID = "CFLCARDCODE";
                oCFLCreationParams01.MultiSelection = false;
                oCFL01 = oCFLs01.Add(oCFLCreationParams01);

                oEdit01.ChooseFromListUID = "CFLCARDCODE";
                oEdit01.ChooseFromListAlias = "CardCode";

                oCons = oCFL01.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "C";
                oCFL01.SetConditions(oCons);

                oEdit02 = oForm.Items.Item("ShipTo").Specific;
                oCFLs02 = oForm.ChooseFromLists;
                oCFLCreationParams02 = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams02.ObjectType = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                oCFLCreationParams02.UniqueID = "CFLSHIPCODE";
                oCFLCreationParams02.MultiSelection = false;
                oCFL02 = oCFLs02.Add(oCFLCreationParams02);

                oEdit02.ChooseFromListUID = "CFLSHIPCODE";
                oEdit02.ChooseFromListAlias = "CardCode";
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oCFL01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL01);
                }

                if (oCFL02 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL02);
                }

                if (oCons != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCons);
                }

                if (oCon != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCon);
                }

                if (oCFLs01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs01);
                }

                if (oCFLs02 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs02);
                }

                if (oCFLCreationParams01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams01);
                }

                if (oCFLCreationParams02 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams02);
                }

                if (oEdit01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit01);
                }

                if (oEdit02 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit02);
                }
            }
        }

        /// <summary>
        /// 메뉴활성화
        /// </summary>
        private void PS_SD091_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1293", true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PS_SD091_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_SD091_EnableFormItem();
                    PS_SD091_AddMatrixRow(0, true);
                }
                else
                {
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각모드에따른 아이템설정
        /// </summary>
        private void PS_SD091_EnableFormItem()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.EnableMenu("1293", true);

                    string userBPLID = dataHelpClass.User_BPLID();

                    if (userBPLID == "1" || userBPLID == "4")
                    {
                        oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    
                    oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("Btn01").Enabled = false;
                    PS_SD091_SetDocEntry();
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.EnableMenu("1293", true);
                    oForm.Items.Item("Btn01").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.EnableMenu("1281", true);
                    oForm.EnableMenu("1282", true);
                    oForm.EnableMenu("1293", false);
                    oForm.Items.Item("Btn01").Enabled = true;
                }

                oMat01.AutoResizeColumns();
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_SD091_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD091'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 Row추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_SD091_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_SD091L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD091L.Offset = oRow;
                oDS_PS_SD091L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        /// 필수입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_SD091_CheckDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            int i;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value))
                {
                    errMessage = "사업장은 필수입니다.";
                    oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                {
                    errMessage = "요청자는 필수입니다.";
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errMessage = "요청일자는 필수입니다.";
                    oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("OutWhCd").Specific.Value))
                {
                    errMessage = "출고창고는 필수입니다.";
                    oForm.Items.Item("OutWhCd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("InWhCd").Specific.Value))
                {
                    errMessage = "입고창고는 필수입니다.";
                    oForm.Items.Item("InWhCd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemName").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "품목은 필수입니다.";
                        oMat01.Columns.Item("ItemName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    if (Convert.ToInt32(oMat01.Columns.Item("Qty").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "수량은 필수입니다.";
                        oMat01.Columns.Item("ItemName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                }

                oDS_PS_SD091L.RemoveRecord(oDS_PS_SD091L.Size - 1);
                oMat01.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_SD091_SetDocEntry();
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 리포트 출력
        /// </summary>
        [STAThread]
        private void PS_SD091_PrintReport()
        {
            string WinTitle;
            string ReportName;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                WinTitle = "[PS_SD045] 이동 요청서";
                ReportName = "PS_SD045.rpt";
                //프로시저 : PS_SD045_01

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.Value)
                };

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
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
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    //Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    //Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_SD091_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD091_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PS_SD091_PrintReport);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_SD091_EnableFormItem();
                                PS_SD091_AddMatrixRow(oMat01.RowCount, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_SD091_EnableFormItem();
                            }
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
            }
        }

        /// <summary>
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ItemCode" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))
                        {
                            PS_SM020 tempForm = new PS_SM020();
                            tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                            PS_SD091_AddMatrixRow(0, true);
                            BubbleEvent = false;
                        }
                    }
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OutWhCd", "");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "InWhCd", "");
                }
                else if (pVal.Before_Action == false)
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                    }
                }
                else if (pVal.Before_Action == false)
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
                        if (pVal.ItemUID == "Mat01")
                        {
                        }
                        else
                        {
                            if (pVal.ItemUID == "BPLId")
                            {
                                if (oForm.Items.Item("BPLId").Specific.Value == "1")
                                {
                                    oDS_PS_SD091H.SetValue("U_OutWhCd", 0, "104");
                                    oDS_PS_SD091H.SetValue("U_InWhCd", 0, "101");
                                }
                                else if (oForm.Items.Item("BPLId").Specific.Value == "4")
                                {
                                    oDS_PS_SD091H.SetValue("U_OutWhCd", 0, "101");
                                    oDS_PS_SD091H.SetValue("U_InWhCd", 0, "104");
                                }
                            }
                        }
                        oForm.Update();
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                }
                else if (pVal.Before_Action == false)
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sQry;
            double SumQty = 0;
            double SumWeight = 0;
            string itemName;
            string itemBSort;
            string itemMSort;
            string unit1;
            string size;
            string itemType;
            string quality;
            string mark;
            string callSize;
            string sbasUnit;

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "ItemCode") //품목 이름 Query
                            {
                                if ((pVal.Row == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_SD091_AddMatrixRow(pVal.Row, false);
                                    oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }

                                oMat01.FlushToDataSource();

                                sQry = "  Select  ItemName,";
                                sQry += "         U_ItmBsort,";
                                sQry += "         U_ItmMsort,";
                                sQry += "         U_Unit1,";
                                sQry += "         U_Size,";
                                sQry += "         U_ItemType,";
                                sQry += "         U_Quality,";
                                sQry += "         U_Mark,";
                                sQry += "         U_CallSize,";
                                sQry += "         U_SbasUnit";
                                sQry += " From    [OITM]";
                                sQry += " Where   ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";

                                oRecordSet.DoQuery(sQry);

                                itemName = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                                itemBSort = oRecordSet.Fields.Item(1).Value.ToString().Trim();
                                itemMSort = oRecordSet.Fields.Item(2).Value.ToString().Trim();
                                unit1 = oRecordSet.Fields.Item(3).Value.ToString().Trim();
                                size = oRecordSet.Fields.Item(4).Value.ToString().Trim();
                                itemType = oRecordSet.Fields.Item(5).Value.ToString().Trim();
                                quality = oRecordSet.Fields.Item(6).Value.ToString().Trim();
                                mark = oRecordSet.Fields.Item(7).Value.ToString().Trim();
                                callSize = oRecordSet.Fields.Item(8).Value.ToString().Trim();
                                sbasUnit = oRecordSet.Fields.Item(9).Value.ToString().Trim();

                                oDS_PS_SD091L.SetValue("U_ItemName", pVal.Row - 1, itemName);
                                oDS_PS_SD091L.SetValue("U_Unit1", pVal.Row - 1, unit1);
                                oDS_PS_SD091L.SetValue("U_Size", pVal.Row - 1, size);
                                oDS_PS_SD091L.SetValue("U_CallSize", pVal.Row - 1, callSize);

                                //품목 대분류
                                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                sQry = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code = '" + itemBSort + "'";
                                oRecordSet.DoQuery(sQry);
                                oDS_PS_SD091L.SetValue("U_ItmBsort", pVal.Row - 1, oRecordSet.Fields.Item(1).Value);
                                
                                //품목 중분류
                                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                sQry = "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_Code = '" + itemMSort + "'";
                                oRecordSet.DoQuery(sQry);
                                oDS_PS_SD091L.SetValue("U_ItmMsort", pVal.Row - 1, oRecordSet.Fields.Item(1).Value);
                                
                                //형태타입
                                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                sQry = "SELECT Code, Name FROM [@PSH_SHAPE] WHERE Code = '" + itemType + "'";
                                oRecordSet.DoQuery(sQry);
                                oDS_PS_SD091L.SetValue("U_ItemType", pVal.Row - 1, oRecordSet.Fields.Item(1).Value);

                                //질별
                                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                sQry = "SELECT Code, Name FROM [@PSH_QUALITY] WHERE Code = '" + quality + "'";
                                oRecordSet.DoQuery(sQry);
                                oDS_PS_SD091L.SetValue("U_Quality", pVal.Row - 1, oRecordSet.Fields.Item(1).Value);

                                //인증기호
                                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                sQry = "SELECT Code, Name FROM [@PSH_MARK] WHERE Code = '" + mark + "'";
                                oRecordSet.DoQuery(sQry);
                                oDS_PS_SD091L.SetValue("U_Mark", pVal.Row - 1, oRecordSet.Fields.Item(1).Value);

                                //판매기준단위
                                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                sQry = "SELECT Code, Name FROM [@PSH_UOMORG] WHERE Code = '" + sbasUnit + "'";
                                oRecordSet.DoQuery(sQry);
                                oDS_PS_SD091L.SetValue("U_SbasUnit", pVal.Row - 1, oRecordSet.Fields.Item(1).Value);
                            }
                            else if (pVal.ColUID == "Qty")
                            {
                                if (oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value < 0)
                                {
                                    oDS_PS_SD091L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
                                }
                                else
                                {
                                    oDS_PS_SD091L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                            }
                            else
                            {
                                oDS_PS_SD091L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }

                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                                {
                                    SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                                }
                                SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                            }
                            oForm.Items.Item("SumQty").Specific.Value = SumQty;
                            oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_SD091H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "CntcCode")
                            {
                                oDS_PS_SD091H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                sQry = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                            }
                            else
                            {
                                oDS_PS_SD091H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oForm.Update();
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            double SumQty = 0;
            double SumWeight = 0;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                        {
                            SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                        }

                        SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                    }

                    oForm.Items.Item("SumQty").Specific.Value = SumQty;
                    oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

                    PS_SD091_EnableFormItem();
                    PS_SD091_AddMatrixRow(oMat01.VisualRowCount, false);
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
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD091H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD091L);
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
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            SAPbouiCOM.DataTable oDataTable01 = null;
            SAPbouiCOM.DataTable oDataTable02 = null;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "CardCode")
                    {
                        oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당
                        if (oDataTable01 != null)
                        {
                            oDS_PS_SD091H.SetValue("U_CardCode", 0, oDataTable01.Columns.Item("CardCode").Cells.Item(0).Value);
                            oDS_PS_SD091H.SetValue("U_CardName", 0, oDataTable01.Columns.Item("CardName").Cells.Item(0).Value);
                            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                                
                        }
                    }
                    else if (pVal.ItemUID == "ShipTo")
                    {
                        oDataTable02 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당
                        if (oDataTable02 != null)
                        {
                            oDS_PS_SD091H.SetValue("U_ShipTo", 0, oDataTable02.Columns.Item("CardCode").Cells.Item(0).Value);
                            oDS_PS_SD091H.SetValue("U_ShipNm", 0, oDataTable02.Columns.Item("CardName").Cells.Item(0).Value);
                            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }   
                        }
                    }

                    oForm.Update();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oDataTable01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable01);
                }

                if (oDataTable02 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable02);
                }
            }
        }

        /// <summary>
        /// 행삭제 체크 메소드(Raise_FormMenuEvent 에서 사용)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i;
            double SumQty = 0;
            double SumWeight = 0;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_SD091L.RemoveRecord(oDS_PS_SD091L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_SD091_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_SD091L.GetValue("U_ItemCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_SD091_AddMatrixRow(oMat01.RowCount, false);
                            }
                        }

                        for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                            {
                                SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                            }
                            SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);

                        }
                        oForm.Items.Item("SumQty").Specific.Value = SumQty;
                        oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_SD091_EnableFormItem();
                            break;
                        case "1282": //추가
                            PS_SD091_EnableFormItem();
                            PS_SD091_AddMatrixRow(0, true);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            //PS_SD091_EnableFormItem();
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
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
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
        /// RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                }
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                        break;
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

using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Drawing;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 통합수불부
	/// </summary>
	internal class PS_CO605 : PSH_BaseClass
	{
		private string oFormUniqueID;
		//public SAPbouiCOM.Form oForm01;
		private SAPbouiCOM.Grid oGrid01;

		private SAPbouiCOM.DataTable oDS_PS_CO605A;

		//public SAPbouiCOM.Form oBaseForm01; //부모폼
		//public string oBaseItemUID01;
		//public string oBaseColUID01;
		//public int oBaseColRow01;
		//public string oBaseTradeType01;
		//public string oBaseItmBsort01;
			
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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO605.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO605_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO605");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

				oForm.Freeze(true);
                PS_CO605_CreateItems();
                PS_CO605_ComboBox_Setting();
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
        private void PS_CO605_CreateItems()
        {
            SAPbouiCOM.CheckBox oChkBox = null;

            try
            {
                oForm.Freeze(true);
                
                oGrid01 = oForm.Items.Item("Grid01").Specific;
                //oGrid01.SelectionMode = ms_NotSupported

                oForm.DataSources.DataTables.Add("PS_CO605A");
                oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_CO605A");
                oDS_PS_CO605A = oForm.DataSources.DataTables.Item("PS_CO605A");

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

                //전기일자(Fr)
                oForm.DataSources.UserDataSources.Add("StrDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("StrDate").Specific.DataBind.SetBound(true, "", "StrDate");
                oForm.DataSources.UserDataSources.Item("StrDate").Value = DateTime.Now.ToString("yyyyMM01"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM01");

                //전기일자(To)
                oForm.DataSources.UserDataSources.Add("EndDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("EndDate").Specific.DataBind.SetBound(true, "", "EndDate");
                oForm.DataSources.UserDataSources.Item("EndDate").Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");

                //재고계정
                oForm.DataSources.UserDataSources.Add("AcctCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("AcctCode").Specific.DataBind.SetBound(true, "", "AcctCode");

                //창고
                oForm.DataSources.UserDataSources.Add("WhsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("WhsCode").Specific.DataBind.SetBound(true, "", "WhsCode");

                //대분류
                oForm.DataSources.UserDataSources.Add("ItmBSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmBSort").Specific.DataBind.SetBound(true, "", "ItmBSort");

                //중분류
                oForm.DataSources.UserDataSources.Add("ItmMSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmMSort").Specific.DataBind.SetBound(true, "", "ItmMSort");

                //출력구분
                oForm.DataSources.UserDataSources.Add("Gubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Gubun").Specific.DataBind.SetBound(true, "", "Gubun");

                //체크박스 처리
                oForm.DataSources.UserDataSources.Add("Check01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

                oChkBox = oForm.Items.Item("Check01").Specific;
                oChkBox.ValOn = "Y";
                oChkBox.ValOff = "N";
                oChkBox.DataBind.SetBound(true, "", "Check01");

                oForm.DataSources.UserDataSources.Item("Check01").Value = "N"; //미체크로 값을 주고 폼을 로드

                oForm.DataSources.UserDataSources.Add("Check02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

                oChkBox = oForm.Items.Item("Check02").Specific;
                oChkBox.ValOn = "Y";
                oChkBox.ValOff = "N";
                oChkBox.DataBind.SetBound(true, "", "Check02");

                oForm.DataSources.UserDataSources.Item("Check02").Value = "N"; //미체크로 값을 주고 폼을 로드

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oChkBox);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_CO605_ComboBox_Setting()
        {
        
            SAPbouiCOM.ComboBox oCombo = null;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                //콤보에 기본값설정

                //사업장
                oCombo = oForm.Items.Item("BPLID").Specific;
                oCombo.ValidValues.Add("%", "전체");
                oCombo.ValidValues.Add("1", "창원사업장");
                oCombo.ValidValues.Add("2", "부산사업장");
                oCombo.ValidValues.Add("6", "안강+울산사업장");
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //재고계정
                oCombo = oForm.Items.Item("AcctCode").Specific;
                oCombo.ValidValues.Add("11506100", "원재료");
                oCombo.ValidValues.Add("11502100", "제품");
                oCombo.ValidValues.Add("11501100", "상품");
                oCombo.ValidValues.Add("11507100", "저장품");
                oCombo.ValidValues.Add("11503100", "재공품");
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //창고
                oCombo = oForm.Items.Item("WhsCode").Specific;
                sQry = "SELECT WhsCode, WhsName From OWHS";
                oRecordSet01.DoQuery(sQry);
                oCombo.ValidValues.Add("000", "전체");
                while (!oRecordSet01.EoF)
                {
                    oCombo.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //대분류
                oCombo = oForm.Items.Item("ItmBSort").Specific;
                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code";
                oRecordSet01.DoQuery(sQry);
                oCombo.ValidValues.Add("001", "전체");
                while (!oRecordSet01.EoF)
                {
                    oCombo.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //중분류
                oCombo = oForm.Items.Item("ItmMSort").Specific;
                sQry = "SELECT U_Code,U_CodeName FROM [@PSH_ITMMSORT] Order by U_Code";
                oRecordSet01.DoQuery(sQry);

                if (oForm.Items.Item("ItmMSort").Specific.ValidValues.Count == 0)
                {
                    oCombo.ValidValues.Add("00001", "전체");
                }

                while (!oRecordSet01.EoF)
                {
                    oCombo.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //출력구분
                oCombo = oForm.Items.Item("Gubun").Specific;
                oCombo.ValidValues.Add("1", "개별");
                oCombo.ValidValues.Add("2", "집계");
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_CO605_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //각모드에따른 아이템설정
                    //PS_CO605_FormClear //UDO방식
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    //각모드에따른 아이템설정
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    //각모드에따른 아이템설정
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
        /// 매트릭스 행 추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_CO605_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                //if (RowIserted = false) //행추가여부
                //{
                //    oDS_PS_CO605L.InsertRecord(oRow);
                //}
                    
                //oMat01.AddRow();
                //oDS_PS_CO605L.Offset = oRow;
                //oDS_PS_CO605L.setValue("U_LineNum", oRow, oRow + 1);
                //oMat01.LoadFromDataSource();
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

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                    //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
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
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
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
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
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
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
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
            try
            {
                if (pVal.Before_Action == true)
                {
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
            try
            {
                if (pVal.Before_Action == true)
                {
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
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
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
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    //원본 소스(VB6.0 주석처리되어 있음)
                    //if(pVal.ItemUID == "Code")
                    //{
                    //    dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY001A", "Code", "", 0, "", "", "");
                    //}
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
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

















        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	switch (pval.EventType) {
        //		case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //			////1
        //			Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //			////2
        //			Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //			////5
        //			Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CLICK:
        //			////6
        //			Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //			////7
        //			Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //			////8
        //			Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //			////10
        //			Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //			////11
        //			Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //			////18
        //			break;
        //		////et_FORM_ACTIVATE
        //		case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //			////19
        //			break;
        //		////et_FORM_DEACTIVATE
        //		case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //			////20
        //			Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //			////27
        //			Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //			////3
        //			Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //			////4
        //			break;
        //		////et_LOST_FOCUS
        //		case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //			////17
        //			Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
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
        //				break;
        //			////Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
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
        //				break;
        //			////Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
        //			case "1281":
        //				//찾기
        //				break;
        //			////Call PS_CO605_FormItemEnabled '//UDO방식
        //			case "1282":
        //				//추가
        //				break;
        //			////Call PS_CO605_FormItemEnabled '//UDO방식
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {
        //		//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
        //		//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
        //		//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        //		//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
        //		//            MenuCreationParams01.uniqueID = "MenuUID"
        //		//            MenuCreationParams01.String = "메뉴명"
        //		//            MenuCreationParams01.Enabled = True
        //		//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
        //		//        End If
        //	} else if (pval.BeforeAction == false) {
        //		//        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
        //		//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
        //		//        End If
        //	}
        //	if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02") {
        //		if (pval.Row > 0) {
        //			oLastItemUID01 = pval.ItemUID;
        //			oLastColUID01 = pval.ColUID;
        //			oLastColRow01 = pval.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pval.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	if (pval.BeforeAction == true) {

        //		if (pval.ItemUID == "BtnSearch") {
        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {

        //				if (PS_CO605_DataValidCheck() == false) {

        //					BubbleEvent = false;
        //					return;

        //				} else {

        //					PS_CO605_MTX01();

        //				}

        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}

        //		} else if (pval.ItemUID == "BtnPrt") {

        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				PS_CO605_Print_Report01();
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}

        //		}

        //		//        If pval.ItemUID = "1" Then
        //		//            If oForm01.Mode = fm_ADD_MODE Then
        //		//                If PS_CO605_DataValidCheck = False Then
        //		//                    BubbleEvent = False
        //		//                    Exit Sub
        //		//                End If
        //		//                '//해야할일 작업
        //		//            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        //		//            ElseIf oForm01.Mode = fm_OK_MODE Then
        //		//            End If
        //		//        End If
        //	} else if (pval.BeforeAction == false) {

        //		if (pval.ItemUID == "PS_CO605") {
        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		//        If pval.ItemUID = "1" Then
        //		//            If oForm01.Mode = fm_ADD_MODE Then
        //		//                If pval.ActionSuccess = True Then
        //		//                    Call PS_CO605_FormItemEnabled
        //		//                    Call PS_CO605_FormClear '//UDO방식일때
        //		//                    Call PS_CO605_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
        //		//                End If
        //		//            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        //		//            ElseIf oForm01.Mode = fm_OK_MODE Then
        //		//                If pval.ActionSuccess = True Then
        //		//                    Call PS_CO605_FormItemEnabled
        //		//                End If
        //		//            End If
        //		//        End If
        //	}
        //	return;
        //	Raise_EVENT_ITEM_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {

        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "CardCode", "");
        //		//거래처
        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "ItemCode", "");
        //		//작번

        //	} else if (pval.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_KEY_DOWN_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	oForm01.Freeze(true);

        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		PS_CO605_FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);
        //	}

        //	oForm01.Freeze(false);

        //	return;
        //	Raise_EVENT_COMBO_SELECT_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CLICK
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {
        //		if (pval.ItemUID == "Grid01") {
        //			if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (pval.Row > 0) {

        //				}
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //	} else if (pval.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
        //private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {
        //		if (pval.ItemUID == "Grid01") {
        //			if (pval.Row == -1) {
        //				//                oGrid01.Columns(pval.ColUID).TitleObject.Sortable = True

        //			} else {
        //				if (oGrid01.Rows.SelectedRows.Count > 0) {

        //					//Call PS_CO605_GetDetail

        //					//                    Call PS_CO605_SetBaseForm '//부모폼에입력
        //					//                    If Trim(oForm01.DataSources.UserDataSources("Check01").VALUE) = "N" Then
        //					//                        Call oForm01.Close
        //					//                    End If
        //				} else {
        //					BubbleEvent = false;
        //				}
        //			}
        //		}
        //	} else if (pval.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_DOUBLE_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
        //private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_VALIDATE
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	oForm01.Freeze(true);

        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {

        //		if (pval.ItemChanged == true) {
        //			PS_CO605_FlushToItemValue(pval.ItemUID);
        //		}

        //	}

        //	oForm01.Freeze(false);

        //	return;
        //	Raise_EVENT_VALIDATE_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		PS_CO605_FormItemEnabled();
        //		////Call PS_CO605_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RESIZE
        //private void Raise_EVENT_RESIZE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		PS_CO605_FormResize();
        //	}
        //	return;
        //	Raise_EVENT_RESIZE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	SAPbouiCOM.DataTable oDataTable01 = null;
        //	if (pval.BeforeAction == true) {

        //	} else if (pval.BeforeAction == false) {
        //		//If (pval.ItemUID = "ItemCode") Then
        //		//   Set oDataTable01 = pval.SelectedObjects
        //		//    If oDataTable01 Is Nothing Then
        //		//    Else
        //		//  oForm01.DataSources.UserDataSources("ItemCode").VALUE = oDataTable01.Columns(0).Cells(0).VALUE
        //		//     '  oForm01.DataSources.UserDataSources("ItemName").VALUE = oDataTable01.Columns(1).Cells(0).VALUE
        //		//   End If
        //		// End If
        //		oForm01.Update();
        //	}
        //	//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDataTable01 = null;
        //	return;
        //	Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02") {
        //		if (pval.Row > 0) {
        //			oLastItemUID01 = pval.ItemUID;
        //			oLastColUID01 = pval.ColUID;
        //			oLastColRow01 = pval.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pval.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}

        //	return;
        //	Raise_EVENT_GOT_FOCUS_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true) {
        //	} else if (pval.BeforeAction == false) {
        //		SubMain.RemoveForms(oFormUniqueID01);
        //		//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oForm01 = null;
        //		//UPGRADE_NOTE: oGrid01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oGrid01 = null;
        //	}
        //	return;
        //	Raise_EVENT_FORM_UNLOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        //			//        For i = 1 To oMat01.VisualRowCount
        //			//            oMat01.Columns("COL01").Cells(i).Specific.Value = i
        //			//        Next i
        //			//        oMat01.FlushToDataSource
        //			//        Call oDS_PS_CO605L.RemoveRecord(oDS_PS_CO605L.Size - 1)
        //			//        oMat01.LoadFromDataSource
        //			//        If oMat01.RowCount = 0 Then
        //			//            Call PS_CO605_AddMatrixRow(0)
        //			//        Else
        //			//            If Trim(oDS_SM020L.GetValue("U_기준컬럼", oMat01.RowCount - 1)) <> "" Then
        //			//                Call PS_CO605_AddMatrixRow(oMat01.RowCount)
        //			//            End If
        //			//        End If
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion














        #region PS_CO605_DataValidCheck
        //public bool PS_CO605_DataValidCheck()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;

        //	functionReturnValue = true;
        //	return functionReturnValue;
        //	PS_CO605_DataValidCheck_Error:

        //	//    If oForm01.Items("WorkGbn").Specific.Selected.VALUE = "%" Then
        //	//        Sbo_Application.SetStatusBarMessage "작업구분은 필수입니다.", bmt_Short, True
        //	//        oForm01.Items("WorkGbn").Click ct_Regular
        //	//        PS_CO605_DataValidCheck = False
        //	//        Exit Function
        //	//    End If

        //	//    If oMat01.VisualRowCount = 0 Then
        //	//        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        //	//        PS_CO605_DataValidCheck = False
        //	//        Exit Function
        //	//    End If
        //	//    For i = 1 To oMat01.VisualRowCount
        //	//        If (oMat01.Columns("ItemName").Cells(i).Specific.Value = "") Then
        //	//            Sbo_Application.SetStatusBarMessage "품목은 필수입니다.", bmt_Short, True
        //	//            oMat01.Columns("ItemName").Cells(i).Click ct_Regular
        //	//            PS_CO605_DataValidCheck = False
        //	//            Exit Function
        //	//        End If
        //	//    Next
        //	//    Call oDS_SM020L.RemoveRecord(oDS_SM020L.Size - 1)
        //	//    Call oMat01.LoadFromDataSource
        //	//    Call PS_CO605_FormClear
        //	functionReturnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO605_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_CO605_FlushToItemValue
        //private void PS_CO605_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	short i = 0;
        //	short ErrNum = 0;
        //	string sQry = null;
        //	string ItemCode = null;

        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string OrdNum = null;
        //	string SubNo1 = null;
        //	string SubNo2 = null;

        //	switch (oUID) {

        //		//        Case "CardCode"
        //		//
        //		//            oForm01.Items("CardName").Specific.VALUE = MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'") '거래처
        //		//
        //		//        Case "ItemCode"
        //		//
        //		//            oForm01.Items("ItemName").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'") '작번
        //		//            oForm01.Items("ItemSpec").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'") '규격

        //	}

        //	//    If oUID = "OrdNum" Or oUID = "SubNo1" Or oUID = "SubNo2" Then
        //	//
        //	//        OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE)
        //	//        SubNo1 = oForm01.Items("SubNo1").Specific.VALUE
        //	//        SubNo2 = oForm01.Items("SubNo2").Specific.VALUE
        //	//
        //	//        sQry = "           SELECT   CASE"
        //	//        sQry = sQry & "                 WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        //	//        sQry = sQry & "                 ELSE T0.U_JakMyung"
        //	//        sQry = sQry & "             END AS [ItemName],"
        //	//        sQry = sQry & "             CASE"
        //	//        sQry = sQry & "                 WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)"
        //	//        sQry = sQry & "                 ELSE T0.U_JakSize"
        //	//        sQry = sQry & "             END AS [ItemSpec]"
        //	//        sQry = sQry & " FROM     [@PS_PP020H] AS T0"
        //	//        sQry = sQry & " WHERE   T0.U_JakName = '" & OrdNum & "'"
        //	//        sQry = sQry & "             AND T0.U_SubNo1 = CASE WHEN '" & SubNo1 & "' = '' THEN '00' ELSE '" & SubNo1 & "' END"
        //	//        sQry = sQry & "             AND T0.U_SubNo2 = CASE WHEN '" & SubNo2 & "' = '' THEN '000' ELSE '" & SubNo2 & "' END"
        //	//
        //	//        Call oRecordSet01.DoQuery(sQry)
        //	//
        //	//        oForm01.Items("ItemName").Specific.VALUE = oRecordSet01.Fields("ItemName").VALUE
        //	//        oForm01.Items("ItemSpec").Specific.VALUE = oRecordSet01.Fields("ItemSpec").VALUE
        //	//
        //	//    End If

        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	return;
        //	PS_CO605_FlushToItemValue_Error:

        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;

        //	MDC_Com.MDC_GF_Message(ref "PS_CO605_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");

        //}
        #endregion

        #region PS_CO605_MTX01
        //private void PS_CO605_MTX01()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_CO605_MTX01()
        //	//해당모듈    : PS_CO605
        //	//기능        : 수불부조회
        //	//인수        : 없음
        //	//반환값      : 없음
        //	//특이사항    : 없음
        //	//******************************************************************************
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	////메트릭스에 데이터 로드
        //	oForm01.Freeze(true);

        //	int i = 0;
        //	string Query01 = null;
        //	string Query02 = null;

        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string ItmBsort = null;
        //	string ItmMsort = null;
        //	string BPLId = null;
        //	string StrDate = null;
        //	string EndDate = null;
        //	string SItemCode = null;
        //	string EITemCode = null;
        //	string AcctCode = null;
        //	string WhsCode = null;
        //	string ChkBox = null;
        //	string ChkBox02 = null;
        //	string Gubun = null;

        //	//조회조건문
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ItmBsort = Strings.Trim(oForm01.Items.Item("ItmBSort").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ItmMsort = Strings.Trim(oForm01.Items.Item("ItmMSort").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	BPLId = Strings.Trim(oForm01.Items.Item("BPLID").Specific.Selected.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	StrDate = Strings.Trim(oForm01.Items.Item("StrDate").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	EndDate = Strings.Trim(oForm01.Items.Item("EndDate").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	AcctCode = Strings.Trim(oForm01.Items.Item("AcctCode").Specific.Selected.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	WhsCode = Strings.Trim(oForm01.Items.Item("WhsCode").Specific.Selected.VALUE);
        //	ChkBox = Strings.Trim(oForm01.DataSources.UserDataSources.Item("Check01").Value);
        //	ChkBox02 = Strings.Trim(oForm01.DataSources.UserDataSources.Item("Check02").Value);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Gubun = Strings.Trim(oForm01.Items.Item("Gubun").Specific.Selected.VALUE);

        //	if (string.IsNullOrEmpty(StrDate))
        //		StrDate = "19000101";
        //	if (string.IsNullOrEmpty(EndDate))
        //		EndDate = "21001231";
        //	if (ItmBsort == "001")
        //		ItmBsort = "%";
        //	if (ItmMsort == "00001")
        //		ItmMsort = "%";

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	if (Gubun == "1") {
        //		////수불 개별
        //		if (ChkBox02 == "Y") {
        //			//sQry = "EXEC [PS_MM209_03] '" & BPLId & "','" & StrDate & "','" & EndDate & "','" & AcctCode & "', '" & WhsCode & "', '" & ChkBox & "', '" & ItmBsort & "', '" & ItmMsort & "'"
        //			//포장사업팀용
        //			Query01 = "EXEC [PS_MM209_10] '" + BPLId + "','" + StrDate + "','" + EndDate + "','" + AcctCode + "', '" + WhsCode + "', '" + ChkBox + "', '" + ItmBsort + "', '" + ItmMsort + "'";
        //		} else {
        //			Query01 = "EXEC [PS_MM209_02] '" + BPLId + "','" + StrDate + "','" + EndDate + "','" + AcctCode + "', '" + WhsCode + "', '" + ChkBox + "', '" + ItmBsort + "', '" + ItmMsort + "', 'PS_CO605_02'";

        //		}
        //	} else {
        //		////수불집계
        //		Query01 = "EXEC [PS_MM209_04] '" + BPLId + "','" + StrDate + "','" + EndDate + "','" + AcctCode + "', '" + WhsCode + "', '" + ChkBox + "', '" + ItmBsort + "', '" + ItmMsort + "', 'PS_CO605_02'";
        //	}

        //	oGrid01.DataTable.Clear();
        //	oDS_PS_CO605A.ExecuteQuery(Query01);

        //	//    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

        //	oGrid01.Columns.Item(5).RightJustified = true;
        //	oGrid01.Columns.Item(6).RightJustified = true;
        //	oGrid01.Columns.Item(7).RightJustified = true;
        //	oGrid01.Columns.Item(8).RightJustified = true;
        //	oGrid01.Columns.Item(9).RightJustified = true;
        //	oGrid01.Columns.Item(10).RightJustified = true;
        //	oGrid01.Columns.Item(11).RightJustified = true;
        //	oGrid01.Columns.Item(12).RightJustified = true;
        //	oGrid01.Columns.Item(13).RightJustified = true;
        //	oGrid01.Columns.Item(14).RightJustified = true;
        //	oGrid01.Columns.Item(15).RightJustified = true;
        //	if (Gubun == "1") {
        //		oGrid01.Columns.Item(16).RightJustified = true;
        //	}
        //	//    oGrid01.Columns(17).RightJustified = True
        //	//    oGrid01.Columns(18).RightJustified = True
        //	//    oGrid01.Columns(19).RightJustified = True
        //	//    oGrid01.Columns(20).RightJustified = True
        //	//    oGrid01.Columns(21).RightJustified = True
        //	//    oGrid01.Columns(22).RightJustified = True

        //	//    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
        //	//    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
        //	//    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑

        //	//    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
        //	//    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
        //	//    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
        //	//    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
        //	//    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
        //	//    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강

        //	if (oGrid01.Rows.Count == 0) {
        //		MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
        //		goto PS_CO605_MTX01_Exit;
        //	}

        //	oGrid01.AutoResizeColumns();
        //	oForm01.Update();

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO605_MTX01_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm01.Freeze(false);
        //	return;
        //	PS_CO605_MTX01_Error:

        //	oForm01.Freeze(false);

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO605_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO605_DI_API
        //private bool PS_CO605_DI_API()
        //{
        //	//On Error GoTo PS_CO605_DI_API_Error
        //	//    PS_CO605_DI_API = True
        //	//    Dim i, j As Long
        //	//    Dim oDIObject As SAPbobsCOM.Documents
        //	//    Dim RetVal As Long
        //	//    Dim LineNumCount As Long
        //	//    Dim ResultDocNum As Long
        //	//    If Sbo_Company.InTransaction = True Then
        //	//        Sbo_Company.EndTransaction wf_RollBack
        //	//    End If
        //	//    Sbo_Company.StartTransaction
        //	//
        //	//    ReDim ItemInformation(0)
        //	//    ItemInformationCount = 0
        //	//    For i = 1 To oMat01.VisualRowCount
        //	//        ReDim Preserve ItemInformation(ItemInformationCount)
        //	//        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.Value
        //	//        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.Value
        //	//        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.Value
        //	//        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.Value
        //	//        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.Value
        //	//        ItemInformation(ItemInformationCount).Check = False
        //	//        ItemInformationCount = ItemInformationCount + 1
        //	//    Next
        //	//
        //	//    LineNumCount = 0
        //	//    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
        //	//    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
        //	//    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
        //	//    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.Value, "&&&&-&&-&&")
        //	//    For i = 0 To UBound(ItemInformation)
        //	//        If ItemInformation(i).Check = True Then
        //	//            GoTo Continue_First
        //	//        End If
        //	//        If i <> 0 Then
        //	//            oDIObject.Lines.Add
        //	//        End If
        //	//        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        //	//        oDIObject.Lines.WarehouseCode = Trim(oForm01.Items("WhsCode").Specific.VALUE)
        //	//        oDIObject.Lines.BaseType = "22"
        //	//        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
        //	//        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
        //	//        For j = i To UBound(ItemInformation)
        //	//            If ItemInformation(j).Check = True Then
        //	//                GoTo Continue_Second
        //	//            End If
        //	//            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
        //	//                GoTo Continue_Second
        //	//            End If
        //	//            '//같은것
        //	//            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
        //	//            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).BatchNum
        //	//            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
        //	//            oDIObject.Lines.BatchNumbers.Add
        //	//            ItemInformation(j).PDN1No = LineNumCount
        //	//            ItemInformation(j).Check = True
        //	//Continue_Second:
        //	//        Next
        //	//        LineNumCount = LineNumCount + 1
        //	//Continue_First:
        //	//    Next
        //	//    RetVal = oDIObject.Add
        //	//    If RetVal = 0 Then
        //	//        ResultDocNum = Sbo_Company.GetNewObjectKey
        //	//        For i = 0 To UBound(ItemInformation)
        //	//            Call oDS_PS_CO605L.setValue("U_OPDNNo", i, ResultDocNum)
        //	//            Call oDS_PS_CO605L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
        //	//        Next
        //	//    Else
        //	//        GoTo PS_CO605_DI_API_Error
        //	//    End If
        //	//
        //	//    If Sbo_Company.InTransaction = True Then
        //	//        Sbo_Company.EndTransaction wf_Commit
        //	//    End If
        //	//    oMat01.LoadFromDataSource
        //	//    oMat01.AutoResizeColumns
        //	//
        //	//    Set oDIObject = Nothing
        //	//    Exit Function
        //	//PS_CO605_DI_API_DI_Error:
        //	//    If Sbo_Company.InTransaction = True Then
        //	//        Sbo_Company.EndTransaction wf_RollBack
        //	//    End If
        //	//    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
        //	//    PS_CO605_DI_API = False
        //	//    Set oDIObject = Nothing
        //	//    Exit Function
        //	//PS_CO605_DI_API_Error:
        //	//    Sbo_Application.SetStatusBarMessage "PS_CO605_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
        //	//    PS_CO605_DI_API = False
        //}
        #endregion

        #region PS_CO605_SetBaseForm
        //private void PS_CO605_SetBaseForm()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	string ItemCode01 = null;
        //	SAPbouiCOM.Matrix oBaseMat01 = null;
        //	if (oBaseForm01 == null) {
        //		////DoNothing
        //	} else {

        //	}
        //	return;
        //	PS_CO605_SetBaseForm_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO605_SetBaseForm_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO605_FormResize
        //private void PS_CO605_FormResize()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	//그룹박스 크기 동적 할당
        //	//    oForm01.Items("GrpBox01").Height = oForm01.Items("Grid01").Height + 30
        //	//    oForm01.Items("GrpBox01").Width = oForm01.Items("Grid01").Width + 30

        //	if (oGrid01.Columns.Count > 0) {
        //		oGrid01.AutoResizeColumns();
        //	}

        //	return;
        //	PS_CO605_FormResize_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_CO605_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO605_Print_Report01
        //private void PS_CO605_Print_Report01()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	short i = 0;
        //	short ErrNum = 0;
        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry = null;
        //	string Sub_sQry = null;

        //	string ItmGrp = null;
        //	string ItmBsort = null;
        //	string ItmMsort = null;
        //	string BPLId = null;
        //	string StrDate = null;
        //	string EndDate = null;
        //	string SItemCode = null;
        //	string EITemCode = null;
        //	string AcctCode = null;
        //	string WhsCode = null;
        //	string ChkBox = null;
        //	string ChkBox02 = null;
        //	string Gubun = null;


        //	SAPbobsCOM.Recordset oRecordSet = null;

        //	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	MDC_PS_Common.ConnectODBC();

        //	//// 조회조건문
        //	//    ItmGrp = Trim(oForm01.Items("ItmGrp").Specific.Selected.VALUE)
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ItmBsort = Strings.Trim(oForm01.Items.Item("ItmBSort").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ItmMsort = Strings.Trim(oForm01.Items.Item("ItmMSort").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	BPLId = Strings.Trim(oForm01.Items.Item("BPLID").Specific.Selected.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	StrDate = Strings.Trim(oForm01.Items.Item("StrDate").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	EndDate = Strings.Trim(oForm01.Items.Item("EndDate").Specific.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	AcctCode = Strings.Trim(oForm01.Items.Item("AcctCode").Specific.Selected.VALUE);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	WhsCode = Strings.Trim(oForm01.Items.Item("WhsCode").Specific.Selected.VALUE);
        //	ChkBox = Strings.Trim(oForm01.DataSources.UserDataSources.Item("Check01").Value);
        //	ChkBox02 = Strings.Trim(oForm01.DataSources.UserDataSources.Item("Check02").Value);
        //	//UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Gubun = Strings.Trim(oForm01.Items.Item("Gubun").Specific.Selected.VALUE);
        //	//    SItemCode = Trim(oForm01.Items("SItemCode").Specific.VALUE)
        //	//    EITemCode = Trim(oForm01.Items("EItemCode").Specific.VALUE)

        //	if (string.IsNullOrEmpty(StrDate))
        //		StrDate = "19000101";
        //	if (string.IsNullOrEmpty(EndDate))
        //		EndDate = "21001231";

        //	if (ItmBsort == "001")
        //		ItmBsort = "%";
        //	if (ItmMsort == "00001")
        //		ItmMsort = "%";

        //	//    If BPLId = "0" Then
        //	//        BPLId = "%"
        //	//    ElseIf BPLId = "1" Then
        //	//        BPLId = "%1"
        //	//    ElseIf BPLId = "2" Then
        //	//        BPLId = "%2"
        //	//    ElseIf BPLId = "3" Then
        //	//        BPLId = "%3"
        //	//    ElseIf BPLId = "4" Then
        //	//        BPLId = "%4"
        //	//    ElseIf BPLId = "5" Then
        //	//        BPLId = "%5"
        //	//    End If

        //	/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

        //	if (Gubun == "1") {
        //		WinTitle = "[PS_CO605] 수불명세서";
        //		ReportName = "PS_MM209_01.RPT";
        //	} else if (Gubun == "2") {
        //		WinTitle = "[PS_CO605] 수불명세서(집계)";
        //		ReportName = "PS_MM209_04.RPT";
        //	}
        //	MDC_Globals.gRpt_Formula = new string[6];
        //	MDC_Globals.gRpt_Formula_Value = new string[6];

        //	//// Formula 수식필드

        //	MDC_Globals.gRpt_Formula[1] = "StrDate";
        //	MDC_Globals.gRpt_Formula_Value[1] = (string.IsNullOrEmpty(StrDate) ? "All" : Microsoft.VisualBasic.Compatibility.VB6.Support.Format(StrDate, "0000-00-00"));

        //	MDC_Globals.gRpt_Formula[2] = "EndDate";
        //	MDC_Globals.gRpt_Formula_Value[2] = (string.IsNullOrEmpty(EndDate) ? "All" : Microsoft.VisualBasic.Compatibility.VB6.Support.Format(EndDate, "0000-00-00"));

        //	MDC_Globals.gRpt_Formula[3] = "BPLId";

        //	if (BPLId == "6") {
        //		MDC_Globals.gRpt_Formula_Value[3] = "안강+울산 사업장";
        //	} else if (BPLId == "%") {
        //		MDC_Globals.gRpt_Formula_Value[3] = "통합사업장";
        //	} else {
        //		sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" + BPLId + "'";
        //		oRecordSet.DoQuery(sQry);
        //		//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		MDC_Globals.gRpt_Formula_Value[3] = oRecordSet.Fields.Item(0).Value;
        //	}


        //	MDC_Globals.gRpt_Formula[4] = "AcctCode";
        //	MDC_Globals.gRpt_Formula_Value[4] = AcctCode;

        //	MDC_Globals.gRpt_Formula[5] = "ChkBox";
        //	MDC_Globals.gRpt_Formula_Value[5] = ChkBox;

        //	sQry = "SELECT WhsName From OWHS where WhsCode = '" + WhsCode + "'";
        //	oRecordSet.DoQuery(sQry);
        //	MDC_Globals.gRpt_Formula[5] = "WhsName";
        //	MDC_Globals.gRpt_Formula_Value[5] = (WhsCode == "000" ? "전체" : oRecordSet.Fields.Item(0).Value);
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //	//// SubReport


        //	MDC_Globals.gRpt_SFormula[1, 1] = "";
        //	MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

        //	if (Gubun == "1") {
        //		////수불 개별
        //		/// Procedure 실행"
        //		if (ChkBox02 == "Y") {
        //			//sQry = "EXEC [PS_MM209_03] '" & BPLId & "','" & StrDate & "','" & EndDate & "','" & AcctCode & "', '" & WhsCode & "', '" & ChkBox & "', '" & ItmBsort & "', '" & ItmMsort & "'"
        //			//포장사업팀용
        //			sQry = "EXEC [PS_MM209_10] '" + BPLId + "','" + StrDate + "','" + EndDate + "','" + AcctCode + "', '" + WhsCode + "', '" + ChkBox + "', '" + ItmBsort + "', '" + ItmMsort + "'";
        //		} else {
        //			sQry = "EXEC [PS_MM209_02] '" + BPLId + "','" + StrDate + "','" + EndDate + "','" + AcctCode + "', '" + WhsCode + "', '" + ChkBox + "', '" + ItmBsort + "', '" + ItmMsort + "', 'PS_MM209'";

        //		}
        //	} else {
        //		////수불집계
        //		sQry = "EXEC [PS_MM209_04] '" + BPLId + "','" + StrDate + "','" + EndDate + "','" + AcctCode + "', '" + WhsCode + "', '" + ChkBox + "', '" + ItmBsort + "', '" + ItmMsort + "', 'PS_MM209'";
        //	}

        //	//    oRecordSet.DoQuery sQry
        //	//    If oRecordSet.RecordCount = 0 Then
        //	//        ErrNum = 1
        //	//        GoTo Print_Query_Error
        //	//    End If

        //	/// Action (sub_query가 있을때는 'Y'로...)/
        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") == false) {
        //		goto Print_Query_Error;
        //	}

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet = null;
        //	return;
        //	Print_Query_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet = null;

        //	//    If ErrNum = 1 Then
        //	//        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
        //	//    Else
        //	MDC_Com.MDC_GF_Message(ref "Print_Query_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	//    End If
        //}
        #endregion
    }
}

using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 생산완료등록
	/// </summary>
	internal class PS_PP080 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_PP080H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP080L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private string oDocEntry;
		private string oStatus;
		private string oCanceled;

        //DI API 연동용 내부 클래스
        public class ItemInformation
        {
            public string OrdGbn; //작지구분
            public string PP030HNo;
            public string PP030MNo;
            public string ItemCode;
            public string BatchNum;
            public double Quantity;
            public string WhsCode;
            public int LineNum;
            public string ORDRNo; //수주번호
            public string RDR1No; //수주라인번호
        }

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP080.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP080_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP080");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

                PS_PP080_CreateItems();
                PS_PP080_SetComboBox();
                PS_PP080_CF_ChooseFromList();
                PS_PP080_EnableMenus();
                PS_PP080_SetDocument(oFormDocEntry);
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
        private void PS_PP080_CreateItems()
        {
            try
            {
                oDS_PS_PP080H = oForm.DataSources.DBDataSources.Item("@PS_PP080H");
                oDS_PS_PP080L = oForm.DataSources.DBDataSources.Item("@PS_PP080L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //합계수량
                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");

                //기계공구류 수주 수량 금액
                oForm.DataSources.UserDataSources.Add("SjQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SjQty").Specific.DataBind.SetBound(true, "", "SjQty");

                oForm.DataSources.UserDataSources.Add("SjAmt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SjAmt").Specific.DataBind.SetBound(true, "", "SjAmt");

                //기계공구류 비용
                oForm.DataSources.UserDataSources.Add("Cost01", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("Cost01").Specific.DataBind.SetBound(true, "", "Cost01");
                oForm.DataSources.UserDataSources.Add("Cost02", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("Cost02").Specific.DataBind.SetBound(true, "", "Cost02");
                oForm.DataSources.UserDataSources.Add("Cost03", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("Cost03").Specific.DataBind.SetBound(true, "", "Cost03");
                oForm.DataSources.UserDataSources.Add("Cost04", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("Cost04").Specific.DataBind.SetBound(true, "", "Cost04");
                oForm.DataSources.UserDataSources.Add("Cost05", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("Cost05").Specific.DataBind.SetBound(true, "", "Cost05");
                oForm.DataSources.UserDataSources.Add("Cost06", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("Cost06").Specific.DataBind.SetBound(true, "", "Cost06");
                oForm.DataSources.UserDataSources.Add("CostTot", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("CostTot").Specific.DataBind.SetBound(true, "", "CostTot");

                //기계공구류 실적비용
                oForm.DataSources.UserDataSources.Add("aCost01", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("aCost01").Specific.DataBind.SetBound(true, "", "aCost01");
                oForm.DataSources.UserDataSources.Add("aCost02", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("aCost02").Specific.DataBind.SetBound(true, "", "aCost02");
                oForm.DataSources.UserDataSources.Add("aCost03", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("aCost03").Specific.DataBind.SetBound(true, "", "aCost03");
                oForm.DataSources.UserDataSources.Add("aCost04", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("aCost04").Specific.DataBind.SetBound(true, "", "aCost04");
                oForm.DataSources.UserDataSources.Add("aCost05", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("aCost05").Specific.DataBind.SetBound(true, "", "aCost05");
                oForm.DataSources.UserDataSources.Add("aCost06", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("aCost06").Specific.DataBind.SetBound(true, "", "aCost06");
                oForm.DataSources.UserDataSources.Add("aCostTot", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("aCostTot").Specific.DataBind.SetBound(true, "", "aCostTot");
                oForm.DataSources.UserDataSources.Add("tSaleQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("tSaleQty").Specific.DataBind.SetBound(true, "", "tSaleQty");
                oForm.DataSources.UserDataSources.Add("tSaleAmt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("tSaleAmt").Specific.DataBind.SetBound(true, "", "tSaleAmt");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP080_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code",  "", false, false);
                oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList 설정
        /// </summary>
        private void PS_PP080_CF_ChooseFromList()
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.Column oColumn = null;

            try
            {
                oColumn = oMat01.Columns.Item("WhsCode");
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLCreationParams.ObjectType = "64"; //SAPbouiCOM.BoLinkedObject.lf_Warehouses
                oCFLCreationParams.UniqueID = "CFLWAREHOUSES";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oColumn.ChooseFromListUID = "CFLWAREHOUSES";
                oColumn.ChooseFromListAlias = "WhsCode";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oColumn != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn);
                }
                if (oCFLCreationParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                }
                if (oCFL != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                }
                if (oCFLs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                }
            }
        }

        /// <summary>
        /// 메뉴활성화
        /// </summary>
        private void PS_PP080_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, true, false, false, false, false, false, false);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry">DocEntry</param>
        private void PS_PP080_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP080_FormItemEnabled();
                    PS_PP080_AddMatrixRow(0, true);
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
        /// 각 모드에 따른 아이템설정
        /// </summary>
        private void PS_PP080_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP080_FormClear();
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                    oDS_PS_PP080H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + dataHelpClass.User_MSTCOD() + "'", 0, 1));
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oMat01.Columns.Item("PP030No").Editable = true;
                    oMat01.Columns.Item("PQty").Editable = true;
                    oMat01.Columns.Item("NQty").Editable = true;
                    oMat01.Columns.Item("WhsCode").Editable = true;
                    oMat01.Columns.Item("Cost01").Editable = true;
                    oMat01.Columns.Item("Cost02").Editable = true;
                    oMat01.Columns.Item("Cost03").Editable = true;
                    oMat01.Columns.Item("Cost04").Editable = true;
                    oMat01.Columns.Item("Cost05").Editable = true;
                    oMat01.Columns.Item("Cost06").Editable = true;
                    oMat01.Columns.Item("CostTot").Editable = true;
                    oMat01.Columns.Item("SaleAmt").Editable = true;
                    oMat01.Columns.Item("Check").Editable = false;
                    oForm.Items.Item("SumQty").Specific.Value = 0;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oMat01.Columns.Item("PP030No").Editable = false;
                    oMat01.Columns.Item("PQty").Editable = false;
                    oMat01.Columns.Item("NQty").Editable = false;
                    oMat01.Columns.Item("WhsCode").Editable = false;
                    oMat01.Columns.Item("Cost01").Editable = false;
                    oMat01.Columns.Item("Cost02").Editable = false;
                    oMat01.Columns.Item("Cost03").Editable = false;
                    oMat01.Columns.Item("Cost04").Editable = false;
                    oMat01.Columns.Item("Cost05").Editable = false;
                    oMat01.Columns.Item("Cost06").Editable = false;
                    oMat01.Columns.Item("CostTot").Editable = false;
                    oMat01.Columns.Item("SaleAmt").Editable = false;
                    oMat01.Columns.Item("Check").Editable = false;
                    oForm.Items.Item("SumQty").Specific.Value = 0;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("OrdGbn").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oMat01.Columns.Item("PP030No").Editable = false;
                    oMat01.Columns.Item("PQty").Editable = false;
                    oMat01.Columns.Item("NQty").Editable = false;
                    oMat01.Columns.Item("WhsCode").Editable = false;
                    oMat01.Columns.Item("Cost01").Editable = false;
                    oMat01.Columns.Item("Cost02").Editable = false;
                    oMat01.Columns.Item("Cost03").Editable = false;
                    oMat01.Columns.Item("Cost04").Editable = false;
                    oMat01.Columns.Item("Cost05").Editable = false;
                    oMat01.Columns.Item("Cost06").Editable = false;
                    oMat01.Columns.Item("CostTot").Editable = false;
                    oMat01.Columns.Item("SaleAmt").Editable = false;

                    if (oDS_PS_PP080H.GetValue("CanCeled", 0).ToString().Trim() == "Y")
                    {
                        oMat01.Columns.Item("Check").Editable = false;
                    }
                    else
                    {
                        oMat01.Columns.Item("Check").Editable = true;
                    }
                    
                    if (oDS_PS_PP080H.GetValue("U_OrdGbn", 0).ToString().Trim() == "104") //멀티
                    {
                        oMat01.Columns.Item("BWeight").Visible = false;
                        oMat01.Columns.Item("PWeight").Visible = false;
                        oMat01.Columns.Item("YWeight").Visible = false;
                        oMat01.Columns.Item("NWeight").Visible = false;
                    }
                    else //그외
                    {
                        oMat01.Columns.Item("BWeight").Visible = true;
                        oMat01.Columns.Item("PWeight").Visible = true;
                        oMat01.Columns.Item("YWeight").Visible = true;
                        oMat01.Columns.Item("NWeight").Visible = true;
                    }
                }
                PS_PP080_SetVisibleItem(false);
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
        private void PS_PP080_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP080'", "");
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
        /// 화면 아이템 Visible 설정
        /// </summary>
        /// <param name="visible"></param>
        private void PS_PP080_SetVisibleItem(bool visible)
        {
            try
            {
                oForm.Freeze(true);

                oForm.Items.Item("t_SjQty").Visible = visible;
                oForm.Items.Item("t_Amt").Visible = visible;
                oForm.Items.Item("t_Cost01").Visible = visible;
                oForm.Items.Item("t_Cost02").Visible = visible;
                oForm.Items.Item("t_Cost03").Visible = visible;
                oForm.Items.Item("t_Cost04").Visible = visible;
                oForm.Items.Item("t_Cost05").Visible = visible;
                oForm.Items.Item("t_Cost06").Visible = visible;
                oForm.Items.Item("t_CostTot").Visible = visible;

                oForm.Items.Item("SjQty").Visible = visible;
                oForm.Items.Item("SjAmt").Visible = visible;
                oForm.Items.Item("Cost01").Visible = visible;
                oForm.Items.Item("Cost02").Visible = visible;
                oForm.Items.Item("Cost03").Visible = visible;
                oForm.Items.Item("Cost04").Visible = visible;
                oForm.Items.Item("Cost05").Visible = visible;
                oForm.Items.Item("Cost06").Visible = visible;
                oForm.Items.Item("CostTot").Visible = visible;

                oForm.Items.Item("t_SaleQty").Visible = visible;
                oForm.Items.Item("t_mAmt").Visible = visible;
                oForm.Items.Item("t_aCost01").Visible = visible;
                oForm.Items.Item("t_aCost02").Visible = visible;
                oForm.Items.Item("t_aCost03").Visible = visible;
                oForm.Items.Item("t_aCost04").Visible = visible;
                oForm.Items.Item("t_aCost05").Visible = visible;
                oForm.Items.Item("t_aCost06").Visible = visible;
                oForm.Items.Item("t_aCostTot").Visible = visible;

                oForm.Items.Item("tSaleQty").Visible = visible;
                oForm.Items.Item("tSaleAmt").Visible = visible;
                oForm.Items.Item("aCost01").Visible = visible;
                oForm.Items.Item("aCost02").Visible = visible;
                oForm.Items.Item("aCost03").Visible = visible;
                oForm.Items.Item("aCost04").Visible = visible;
                oForm.Items.Item("aCost05").Visible = visible;
                oForm.Items.Item("aCost06").Visible = visible;
                oForm.Items.Item("aCostTot").Visible = visible;
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
        /// 메트릭스 Row추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_PP080_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_PP080L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP080L.Offset = oRow;
                oDS_PS_PP080L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private bool PS_PP080_DataValidCheck()
        {
            bool returnValue = false;
            string sQry;
            double RDR1Qty;
            double PP080LQty;
            int i;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP080_FormClear();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errMessage = "완료일자는 필수입니다.";
                    oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 완료일자를 확인하고, 회계부서로 문의하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                {
                    errMessage = "담당자는 필수입니다.";
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (oMat01.VisualRowCount <= 1)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("PP030No").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "작업지시문서는 필수입니다.";
                        oMat01.Columns.Item("PP030No").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "창고코드는 필수입니다.";
                        oMat01.Columns.Item("WhsCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "품목코드는 필수입니다.";
                        oMat01.Columns.Item("ItemCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "생산수량은 필수입니다.";
                        oMat01.Columns.Item("PQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    
                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "106") //수주수량 VS 생산수량 체크(기계)
                    {
                        if (Convert.ToDouble(oMat01.Columns.Item("BQty").Cells.Item(i).Specific.Value) < Convert.ToDouble(oMat01.Columns.Item("YQty").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "수주잔량수량보다 생산수량이 많습니다. 확인바랍니다.";
                            oMat01.Columns.Item("YQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                        
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" && oForm.Items.Item("BPLId").Specific.Selected.Value == "2")
                        {
                            sQry = "Select U_ItmMSort From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);

                            if (oRecordSet01.Fields.Item(0).Value.ToString().Trim() == "10502" || oRecordSet01.Fields.Item(0).Value.ToString().Trim() == "10503" || oRecordSet01.Fields.Item(0).Value.ToString().Trim() == "10504")
                            {
                                sQry = "  Select    Quantity";
                                sQry += " From      RDR1";
                                sQry += " Where     DocEntry = '" + oMat01.Columns.Item("ORDRNo").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                sQry += "           And LineNum = '" + oMat01.Columns.Item("RDR1No").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);

                                RDR1Qty = Convert.ToDouble(oRecordSet01.Fields.Item(0).Value); //수주수량

                                sQry = "  Select    ISNULL(Sum(a.U_PQty),0)";
                                sQry += " From      [@PS_PP080L] a";
                                sQry += "           Inner Join";
                                sQry += "           [@PS_PP080H] b";
                                sQry += "               On a.DocEntry = b.DocEntry ";
                                sQry += " Where     a.U_ORDRNo = '" + oMat01.Columns.Item("ORDRNo").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                sQry += "           And a.U_RDR1No = '" + oMat01.Columns.Item("RDR1No").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                sQry += "           And ISNULL(a.U_Check, 'N') = 'N'";
                                oRecordSet01.DoQuery(sQry);

                                PP080LQty = Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) + Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value); //생산완료수량(집계)

                                if (RDR1Qty == PP080LQty) //최종 생산완료시(수주량 대비 생산완료수량(집계)이 일치하는 경우)
                                {
                                    //검수입고(원재료품의, 외주제작품의, 가공비품의)가 등록 되지 않으면 생산완료 등록 불가(2012.01.12 송명규 수정)
                                    sQry = "EXEC [PS_PP080_09] '" + oMat01.Columns.Item("PP030HNo").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                    oRecordSet01.DoQuery(sQry);
                                    if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) > 0)
                                    {
                                        if (oRecordSet01.Fields.Item(1).Value == "10")
                                        {
                                            errMessage = "" + i + "번 라인 : 원재료품의가 모두 검수입고 되지 않았습니다. 확인해주세요.";
                                            throw new Exception();
                                        }
                                        else if (oRecordSet01.Fields.Item(1).Value == "30")
                                        {
                                            errMessage = "" + i + "번 라인 : 가공비품의가 모두 검수입고 되지 않았습니다. 확인해주세요.";
                                            throw new Exception();
                                        }
                                        else if (oRecordSet01.Fields.Item(1).Value == "40")
                                        {
                                            errMessage = "" + i + "번 라인 : 외주제작품의가 모두 검수입고 되지 않았습니다. 확인해주세요.";
                                            throw new Exception();
                                        }
                                    }
                                    //검수입고(원재료품의, 외주제작품의, 가공비품의)가 등록 되지 않으면 생산완료 등록 불가(2012.01.12 송명규 수정)

                                    //검사여부등록 체크_S(2015.08.28 송명규 수정)
                                    sQry = "EXEC [PS_PP080_10] '" + oMat01.Columns.Item("OrdNum").Cells.Item(i).Specific.Value.ToString().Trim() + "','" + oForm.Items.Item("DocDate").Specific.Value + "'";
                                    oRecordSet01.DoQuery(sQry);
                                    
                                    if (oRecordSet01.Fields.Item("ReturnValue").Value == "0") //검사등록되지 않음
                                    {
                                        errMessage = "" + i + "번 라인 : 검사등록이 되지 않았습니다. 확인해주세요.";
                                        throw new Exception();
                                    }
                                    else if (oRecordSet01.Fields.Item("ReturnValue").Value == "1")
                                    {
                                        errMessage = "" + i + "번 라인 : 검사등록일자(" + oRecordSet01.Fields.Item("ChkDate").Value + ")보다 생산완료일자가 빠릅니다. 생산완료일자를 확인해주세요.";
                                        throw new Exception();
                                    }
                                    //검사여부등록 체크_E(2015.08.28 송명규 수정)

                                    //외주제작품의 일자 체크_S(2016.02.24 송명규 수정)
                                    if (PS_PP080_CheckDate(oMat01.Columns.Item("PP030No").Cells.Item(i).Specific.Value) == false)
                                    {
                                        errMessage = i + "행 [" + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value + "]의 생산완료일은 검수입고일과 같거나 늦어야합니다. 확인하십시오." + (char)13 + "해당 생산완료는 전체가 등록되지 않습니다.";
                                        throw new Exception();
                                    }
                                    //외주제작품의 일자 체크_E(2016.02.24 송명규 수정)
                                }
                                else if (RDR1Qty > PP080LQty) //수주량보다 완료량이 적은 경우(분할 생산완료)
                                {
                                    //※분할생산완료인 경우 해당 구매요청이 1건일 때는 모든 구매요청수량이 검수입고 완료되어야만 생산완료 등록 가능함
                                    sQry = "EXEC [PS_PP080_11] '" + oMat01.Columns.Item("OrdNum").Cells.Item(i).Specific.Value.ToString().Trim() + "'"; //구매요청건수 조회
                                    oRecordSet01.DoQuery(sQry);
                
                                    if (Convert.ToInt16(oRecordSet01.Fields.Item("CNT").Value) == 1) //구매요청 건수가 1건인 경우(1건 넘는경우 체크 회피)
                                    {
                                        //검수입고(원재료품의, 외주제작품의, 가공비품의)가 등록 되지 않으면 생산완료 등록 불가(2012.01.12 송명규 수정)
                                        sQry = "EXEC [PS_PP080_09] '" + oMat01.Columns.Item("PP030HNo").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                        oRecordSet01.DoQuery(sQry);
                                        if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) > 0)
                                        {
                                            if (oRecordSet01.Fields.Item(1).Value == "10")
                                            {
                                                errMessage = "" + i + "번 라인 : 원재료품의가 모두 검수입고 되지 않았습니다. 확인해주세요.";
                                                throw new Exception();
                                            }
                                            else if (oRecordSet01.Fields.Item(1).Value == "30")
                                            {
                                                errMessage = "" + i + "번 라인 : 가공비품의가 모두 검수입고 되지 않았습니다. 확인해주세요.";
                                                throw new Exception();
                                            }
                                            else if (oRecordSet01.Fields.Item(1).Value == "40")
                                            {
                                                errMessage = "" + i + "번 라인 : 외주제작품의가 모두 검수입고 되지 않았습니다. 확인해주세요.";
                                                throw new Exception();
                                            }
                                        }
                                        //검수입고(원재료품의, 외주제작품의, 가공비품의)가 등록 되지 않으면 생산완료 등록 불가(2012.01.12 송명규 수정)

                                        //검사여부등록 체크_S(2015.08.28 송명규 수정)
                                        sQry = "EXEC [PS_PP080_10] '" + oMat01.Columns.Item("OrdNum").Cells.Item(i).Specific.Value.ToString().Trim() + "','" + oForm.Items.Item("DocDate").Specific.Value + "'";
                                        oRecordSet01.DoQuery(sQry);

                                        if (oRecordSet01.Fields.Item("ReturnValue").Value == "0") //검사등록되지 않음
                                        {
                                            errMessage = "" + i + "번 라인 : 검사등록이 되지 않았습니다. 확인해주세요.";
                                            throw new Exception();
                                        }
                                        else if (oRecordSet01.Fields.Item("ReturnValue").Value == "1")
                                        {
                                            errMessage = "" + i + "번 라인 : 검사등록일자(" + oRecordSet01.Fields.Item("ChkDate").Value + ")보다 생산완료일자가 빠릅니다. 생산완료일자를 확인해주세요.";
                                            throw new Exception();
                                        }
                                        //검사여부등록 체크_E(2015.08.28 송명규 수정)

                                        //외주제작품의 일자 체크_S(2016.02.24 송명규 수정)
                                        if (PS_PP080_CheckDate(oMat01.Columns.Item("PP030No").Cells.Item(i).Specific.Value) == false)
                                        {
                                            errMessage = i + "행 [" + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value + "]의 생산완료일은 검수입고일과 같거나 늦어야합니다. 확인하십시오." + (char)13 + "해당 생산완료는 전체가 등록되지 않습니다.";
                                            throw new Exception();
                                        }
                                        //외주제작품의 일자 체크_E(2016.02.24 송명규 수정)
                                    }
                                }
                                else if (RDR1Qty < PP080LQty) //수주량보다 완료량이 많은 경우 무조건 완료를 잡을 수 없게 한다.
                                {
                                    if (Convert.ToDouble(oMat01.Columns.Item("ORDRNo").Cells.Item(i).Specific.Value.ToString().Trim()) == 0 || string.IsNullOrEmpty(oMat01.Columns.Item("ORDRNo").Cells.Item(i).Specific.Value.ToString().Trim()))
                                    {
                                    }
                                    else
                                    {
                                        errMessage = "수주수량보다 생산완료수량이 더 많습니다. 확인해주세요.";
                                        throw new Exception();
                                    }
                                }
                            }
                        }
                    }
                    
                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "102" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //부품,멀티인경우
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "부품,멀티작업은 배치번호가 필수입니다.";
                            throw new Exception();
                        }
                    }
                }

                if (PS_PP080_Validate("검사01") == false)
                {
                    errMessage = " ";
                    throw new Exception();
                }

                oDS_PS_PP080L.RemoveRecord(oDS_PS_PP080L.Size - 1);
                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage == " ")
                {
                }
                else if (errMessage != string.Empty)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 선행프로세스와 일자 비교
        /// </summary>
        /// <param name="pBaseEntry">기준문서번호</param>
        /// <returns>선행프로세스보다 일자가 같거나 느릴 경우(true), 선행프로세스보다 일자가 빠를 경우(false)</returns>
        private bool PS_PP080_CheckDate(string pBaseEntry)
        {
            bool returnValue = false;
            string query;
            string BaseEntry;
            string BaseLine;
            string DocType;
            string CurDocDate;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                BaseEntry = pBaseEntry;
                BaseLine = "";
                DocType = "PS_PP080";
                CurDocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

                query = "EXEC PS_Z_CHECK_DATE '";
                query += BaseEntry + "','";
                query += BaseLine + "','";
                query += DocType + "','";
                query += CurDocDate + "'";

                oRecordSet01.DoQuery(query);

                if (oRecordSet01.Fields.Item("ReturnValue").Value == "False")
                {
                    returnValue = false;
                }
                else
                {
                    returnValue = true;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_PP080_Validate(string ValidateType)
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                if (ValidateType == "검사01")
                {
                    for (int i = 1; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101" 
                         || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "102" 
                         || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104" 
                         || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "107")
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" + oMat01.Columns.Item("PP030No").Cells.Item(i).Specific.Value + "'", 0, 1)) <= 0)
                            {
                                errMessage = "작업지시문서가 존재하지 않습니다.";
                                throw new Exception();
                            }
                        }
                        else if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "106")
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND PS_PP030H.DocEntry = '" + oMat01.Columns.Item("PP030No").Cells.Item(i).Specific.Value + "'", 0, 1)) <= 0)
                            {
                                errMessage = "작업지시문서가 존재하지 않습니다.";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (ValidateType == "수정")
                {
                }
                else if (ValidateType == "행삭제")
                {
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        errMessage = "현재모드는 행삭제가 불가능합니다.";
                        throw new Exception();
                    }
                }
                else if (ValidateType == "취소")
                {
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

            return returnValue;
        }

        /// <summary>
        /// 입고DI(생산입고)
        /// </summary>
        /// <returns></returns>
        private bool PS_PP080_DI_API01()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int i;
            int RetVal;
            string afterDIDocNum;
            string BaseLot;
            SAPbobsCOM.Documents oDIObject = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                PSH_Globals.oCompany.StartTransaction();

                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                List<ItemInformation> itemInfoList = new List<ItemInformation>();

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    ItemInformation itemInfo = new ItemInformation();
                    itemInfo.OrdGbn = oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Selected.Value;
                    itemInfo.PP030HNo = oMat01.Columns.Item("PP030HNo").Cells.Item(i).Specific.Value;
                    itemInfo.PP030MNo = oMat01.Columns.Item("PP030MNo").Cells.Item(i).Specific.Value;
                    itemInfo.ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
                    itemInfo.Quantity = Convert.ToDouble(oMat01.Columns.Item("YQty").Cells.Item(i).Specific.Value);
                    itemInfo.WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
                    itemInfo.BatchNum = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value;
                    
                    itemInfoList.Add(itemInfo);
                }

                oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                oDIObject.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                oDIObject.Comments = "생산완료 (" + oDS_PS_PP080H.GetValue("DocEntry", 0).ToString().Trim() + ") 입고_PS_PP080";
                for (i = 0; i < itemInfoList.Count; i++)
                {
                    if (i != 0)
                    {
                        oDIObject.Lines.Add();
                    }
                    oDIObject.Lines.ItemCode = itemInfoList[i].ItemCode;
                    oDIObject.Lines.WarehouseCode = itemInfoList[i].WhsCode;
                    oDIObject.Lines.Quantity = itemInfoList[i].Quantity;
                    //부품,멀티인경우 배치 선택
                    if (itemInfoList[i].OrdGbn == "102" || itemInfoList[i].OrdGbn == "104")
                    {
                        //배치사용품목이면
                        if (dataHelpClass.GetItem_ManBtchNum(itemInfoList[i].ItemCode) == "Y")
                        {
                            oDIObject.Lines.BatchNumbers.BatchNumber = itemInfoList[i].BatchNum;
                            oDIObject.Lines.BatchNumbers.Quantity = itemInfoList[i].Quantity;
                            oDIObject.Lines.BatchNumbers.Add();
                        }
                    }
                }

                RetVal = oDIObject.Add();

                if (RetVal == 0)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out afterDIDocNum);
                    oForm.Items.Item("OIGNNo").Specific.Value = afterDIDocNum;
                    oDS_PS_PP080H.SetValue("U_OIGNNo", 0, afterDIDocNum);
                    for (i = 0; i < itemInfoList.Count; i++)
                    {
                        //멀티인경우 투입자재의 배치정보를 생성된 재고에 반영
                        if (itemInfoList[i].OrdGbn == "104")
                        {
                            BaseLot = dataHelpClass.GetValue("SELECT U_BatchNum FROM [@PS_PP030L] WHERE DocEntry = '" + itemInfoList[i].PP030HNo + "'", 0, 1);
                            dataHelpClass.DoQuery("UPDATE [OBTN] SET U_BaseLot = '" + BaseLot + "' WHERE ItemCode = '" + itemInfoList[i].ItemCode + "' AND DistNumber = '" + itemInfoList[i].BatchNum + "'");
                        }
                    }
                }
                else
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errCode = "1";
                    throw new Exception();
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (oDIObject != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 출고DI(생산입고 취소)
        /// </summary>
        /// <returns></returns>
        private bool PS_PP080_DI_API03()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int i;
            int RetVal;
            string afterDIDocNum;
            SAPbobsCOM.Documents oDIObject = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                PSH_Globals.oCompany.StartTransaction();

                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                List<ItemInformation> itemInfoList = new List<ItemInformation>();
                
                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OIGENum").Cells.Item(i).Specific.Value.ToString().Trim()))
                    {
                        if (oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked == true)
                        {
                            ItemInformation itemInfo = new ItemInformation
                            {
                                OrdGbn = oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Selected.Value,
                                PP030HNo = oMat01.Columns.Item("PP030HNo").Cells.Item(i).Specific.Value,
                                PP030MNo = oMat01.Columns.Item("PP030MNo").Cells.Item(i).Specific.Value,
                                ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value,
                                Quantity = Convert.ToDouble(oMat01.Columns.Item("YQty").Cells.Item(i).Specific.Value),
                                WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value,
                                BatchNum = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value,
                                LineNum = Convert.ToInt32(oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value),
                                ORDRNo = oMat01.Columns.Item("ORDRNo").Cells.Item(i).Specific.Value,
                                RDR1No = oMat01.Columns.Item("RDR1No").Cells.Item(i).Specific.Value
                            };

                            itemInfoList.Add(itemInfo);
                        }
                    }
                }

                oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                oDIObject.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                oDIObject.UserFields.Fields.Item("U_CancDoc").Value = oForm.Items.Item("OIGNNo").Specific.Value.ToString().Trim();
                oDIObject.Comments = "생산완료 (" + oDS_PS_PP080H.GetValue("DocEntry", 0).ToString().Trim() + ") 취소(출고)_PS_PP080";
                for (i = 0; i < itemInfoList.Count; i++)
                {
                    if (i != 0)
                    {
                        oDIObject.Lines.Add();
                    }
                    oDIObject.Lines.ItemCode = itemInfoList[i].ItemCode;
                    oDIObject.Lines.WarehouseCode = itemInfoList[i].WhsCode;
                    oDIObject.Lines.Quantity = itemInfoList[i].Quantity;
                    //부품,멀티인경우 배치 선택
                    if (itemInfoList[i].OrdGbn == "102" || itemInfoList[i].OrdGbn == "104")
                    {
                        //배치사용품목이면
                        if (dataHelpClass.GetItem_ManBtchNum(itemInfoList[i].ItemCode) == "Y")
                        {
                            oDIObject.Lines.BatchNumbers.BatchNumber = itemInfoList[i].BatchNum;
                            oDIObject.Lines.BatchNumbers.Quantity = itemInfoList[i].Quantity;
                            oDIObject.Lines.BatchNumbers.Add();
                        }
                    }
                }

                RetVal = oDIObject.Add();
                if (RetVal == 0)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out afterDIDocNum);

                    for (i = 0; i < itemInfoList.Count; i++)
                    {
                        dataHelpClass.DoQuery("UPDATE [@PS_PP080L] SET U_OIGENum = '" + afterDIDocNum + "', U_IGE1Num = '" + i + "', U_Check = 'Y' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "' And LineId = '" + itemInfoList[i].LineNum + "'");

                        //휘팅, 부품 실적추가분 취소처리 => 수량을 0으로 처리
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "102")
                        {
                            dataHelpClass.DoQuery("UPDATE [@PS_PP040L] SET U_PQty = 0, U_PWeight = 0, U_YQty = 0, U_YWeight = 0 WHERE DocEntry = '" + oForm.Items.Item("PP040No").Specific.Value + "' And LineId = '" + itemInfoList[i].LineNum + "'");
                        }
                    }
                }
                else
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errCode = "1";
                    throw new Exception();
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                if (oDIObject != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
                }
            }

            return returnValue;
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
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            string errMessage = string.Empty;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Link01")
                    {
                        PS_PP040 tempForm = new PS_PP040();
                        tempForm.LoadForm(oForm.Items.Item("PP040No").Specific.Value);
                        errMessage = " ";
                        throw new Exception();
                    }
                    else if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP080_DataValidCheck() == false)
                            {
                                errMessage = " ";
                                throw new Exception();
                            }

                            //Addon만등록 시 주석_S
                            if (PS_PP080_DI_API01() == false)
                            {
                                PS_PP080_AddMatrixRow(oMat01.VisualRowCount, false);
                                errMessage = " ";
                                throw new Exception();
                            }
                            //Addon만등록 시 주석_E

                            oDocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP080_DataValidCheck() == false)
                            {
                                
                            }
                            if (oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "1" && oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "102")
                            {
                                PSH_Globals.SBO_Application.MessageBox("창원사업장의 부품Item은 갱신할 수 없습니다.");
                                PS_PP080_AddMatrixRow(oMat01.VisualRowCount, false); 
                                errMessage = " ";
                                throw new Exception();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
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
                                dataHelpClass.DoQuery("EXEC PS_PP080_03 '" + oDocEntry + "'");
                                PS_PP080_FormItemEnabled();
                                PS_PP080_AddMatrixRow(0, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_PP080_FormItemEnabled();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    if (errMessage == " ")
                    {
                        BubbleEvent = false;
                    }
                    else
                    {
                        PSH_Globals.SBO_Application.StatusBar.SetText(errMessage);
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
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
            string errMessage = string.Empty;
            SAPbouiCOM.BoStatusBarMessageType messageType = BoStatusBarMessageType.smt_Error;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");

                    if (pVal.ColUID == "PP030No")
                    {
                        if (oForm.Items.Item("BPLId").Specific.Selected.Value == "선택")
                        {
                            errMessage = "사업장은 필수입니다.";
                            messageType = BoStatusBarMessageType.smt_Warning;
                            BubbleEvent = false;
                            throw new Exception();
                        }
                        else if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
                        {
                            errMessage = "작업구분은 필수입니다.";
                            messageType = BoStatusBarMessageType.smt_Warning;
                            BubbleEvent = false;
                            throw new Exception();
                        }
                        else
                        {
                            if (pVal.ColUID == "PP030No" && oForm.Items.Item("BPLId").Specific.Selected.Value == "1" && oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101") //창원 휘팅일때
                            {
                                PS_PP071 tempForm = new PS_PP071();
                                tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row, oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim(), oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim());
                            }
                            else
                            {
                                dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "PP030No");
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.CharPressed == 38) //방향키(↑)
                        {
                            if (pVal.Row > 1 && pVal.Row <= oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row - 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Freeze(false);
                            }
                        }
                        else if (pVal.CharPressed == 40) //방향키(↓)
                        {
                            if (pVal.Row > 0 && pVal.Row < oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Freeze(false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, messageType);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
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
                            if (pVal.ColUID == "특정컬럼")
                            {
                                oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP080_AddMatrixRow(pVal.Row, false);
                                }
                            }
                            else
                            {
                                oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP080H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                            else if (pVal.ItemUID == "BPLId")
                            {
                                oDS_PS_PP080H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP080_AddMatrixRow(0, true);
                            }
                            else if (pVal.ItemUID == "OrdGbn")
                            {
                                oDS_PS_PP080H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP080_AddMatrixRow(0, true);
                                
                                if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "104") //멀티
                                {
                                    oMat01.Columns.Item("BWeight").Visible = false;
                                    oMat01.Columns.Item("PWeight").Visible = false;
                                    oMat01.Columns.Item("YWeight").Visible = false;
                                    oMat01.Columns.Item("NWeight").Visible = false;
                                }
                                else //그외의경우
                                {
                                    oMat01.Columns.Item("BWeight").Visible = true;
                                    oMat01.Columns.Item("PWeight").Visible = true;
                                    oMat01.Columns.Item("YWeight").Visible = true;
                                    oMat01.Columns.Item("NWeight").Visible = true;
                                }
                            }
                            else
                            {
                                oDS_PS_PP080H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                        }

                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oForm.Update();

                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                        else
                        {
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
            string Query01;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        //if (pVal.ColUID != "PQty") //생산수량 클릭시 제외
                        //{
                            if (pVal.Row > 0)
                            {
                                oMat01.SelectRow(pVal.Row, true, false);
                                if (oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "105" || oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "106")
                                {
                                    ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                                    //폼 사용자 필드 visible
                                    PS_PP080_SetVisibleItem(true);

                                    Query01 = "Select Quantity, LineTotal From RDR1 Where DocEntry = '" + oMat01.Columns.Item("ORDRNo").Cells.Item(pVal.Row).Specific.Value + "' And U_LineNum = '" + oMat01.Columns.Item("RDR1No").Cells.Item(pVal.Row).Specific.Value + "'";
                                    RecordSet01.DoQuery(Query01);
                                    if (RecordSet01.RecordCount == 1)
                                    {
                                        oForm.Items.Item("SjQty").Specific.Value = RecordSet01.Fields.Item("Quantity").Value;
                                        oForm.Items.Item("SjAmt").Specific.Value = RecordSet01.Fields.Item("LineTotal").Value;
                                    }
                                    else
                                    {
                                        oForm.Items.Item("SjQty").Specific.Value = 0;
                                        oForm.Items.Item("SjAmt").Specific.Value = 0;
                                    }

                                    Query01 = "EXEC PS_PP080_05 '" + oMat01.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value + "', '" + oMat01.Columns.Item("PP030HNo").Cells.Item(pVal.Row).Specific.Value + "'";
                                    RecordSet01.DoQuery(Query01);

                                    if (RecordSet01.RecordCount == 1)
                                    {
                                        oForm.Items.Item("Cost01").Specific.Value = RecordSet01.Fields.Item("Cost01").Value;
                                        oForm.Items.Item("Cost02").Specific.Value = RecordSet01.Fields.Item("Cost02").Value;
                                        oForm.Items.Item("Cost03").Specific.Value = RecordSet01.Fields.Item("Cost03").Value;
                                        oForm.Items.Item("Cost04").Specific.Value = RecordSet01.Fields.Item("Cost04").Value;
                                        oForm.Items.Item("Cost05").Specific.Value = RecordSet01.Fields.Item("Cost05").Value;
                                        oForm.Items.Item("Cost06").Specific.Value = RecordSet01.Fields.Item("Cost06").Value;
                                        oForm.Items.Item("CostTot").Specific.Value = RecordSet01.Fields.Item("CostTot").Value;
                                        oForm.Items.Item("aCost01").Specific.Value = RecordSet01.Fields.Item("aCost01").Value;
                                        oForm.Items.Item("aCost02").Specific.Value = RecordSet01.Fields.Item("aCost02").Value;
                                        oForm.Items.Item("aCost03").Specific.Value = RecordSet01.Fields.Item("aCost03").Value;
                                        oForm.Items.Item("aCost04").Specific.Value = RecordSet01.Fields.Item("aCost04").Value;
                                        oForm.Items.Item("aCost05").Specific.Value = RecordSet01.Fields.Item("aCost05").Value;
                                        oForm.Items.Item("aCost06").Specific.Value = RecordSet01.Fields.Item("aCost06").Value;
                                        oForm.Items.Item("aCostTot").Specific.Value = RecordSet01.Fields.Item("aCostTot").Value;
                                        oForm.Items.Item("tSaleQty").Specific.Value = RecordSet01.Fields.Item("tSaleQty").Value;
                                        oForm.Items.Item("tSaleAmt").Specific.Value = RecordSet01.Fields.Item("tSaleAmt").Value;
                                    }
                                    else
                                    {
                                        oForm.Items.Item("Cost01").Specific.Value = 0;
                                        oForm.Items.Item("Cost02").Specific.Value = 0;
                                        oForm.Items.Item("Cost03").Specific.Value = 0;
                                        oForm.Items.Item("Cost04").Specific.Value = 0;
                                        oForm.Items.Item("Cost05").Specific.Value = 0;
                                        oForm.Items.Item("Cost06").Specific.Value = 0;
                                        oForm.Items.Item("CostTot").Specific.Value = 0;
                                        oForm.Items.Item("aCost01").Specific.Value = 0;
                                        oForm.Items.Item("aCost02").Specific.Value = 0;
                                        oForm.Items.Item("aCost03").Specific.Value = 0;
                                        oForm.Items.Item("aCost04").Specific.Value = 0;
                                        oForm.Items.Item("aCost05").Specific.Value = 0;
                                        oForm.Items.Item("aCost06").Specific.Value = 0;
                                        oForm.Items.Item("aCostTot").Specific.Value = 0;
                                        oForm.Items.Item("tSaleAmt").Specific.Value = 0;
                                    }
                                }
                            }
                        //}
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
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
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
            string Check = string.Empty;

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.Row == 0 && pVal.ColUID == "Check")
                    {
                        oMat01.FlushToDataSource();
                        if (string.IsNullOrEmpty(oDS_PS_PP080L.GetValue("U_Check", 0).ToString().Trim()) || oDS_PS_PP080L.GetValue("U_Check", 0).ToString().Trim() == "N")
                        {
                            Check = "Y";
                        }
                        else if (oDS_PS_PP080L.GetValue("U_Check", 0).ToString().Trim() == "Y")
                        {
                            Check = "N";
                        }
                        for (int i = 0; i <= oMat01.VisualRowCount - 2; i++)
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_PP080L.GetValue("U_OIGENum", 0).ToString().Trim()))
                            {
                                oDS_PS_PP080L.SetValue("U_Check", i, "Y");
                            }
                            else
                            {
                                oDS_PS_PP080L.SetValue("U_Check", i, Check);
                            }
                        }
                        oMat01.LoadFromDataSource();
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string errMessage = string.Empty;
            int i;
            string query;
            double sumQty = 0;
            double weight;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "PP030No")
                            {
                                oMat01.FlushToDataSource();

                                for (i = 1; i <= oMat01.RowCount; i++)
                                {
                                    if (oMat01.Columns.Item("PP030No").Cells.Item(i).Specific.Value == oMat01.Columns.Item("PP030No").Cells.Item(pVal.Row).Specific.Value && i != pVal.Row) //현재 입력한 값이 이미 입력되어 있는경우
                                    {
                                        errMessage = "이미 입력한 작업지시문서입니다.";
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                        if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                        {
                                            PS_PP080_AddMatrixRow(pVal.Row, false);
                                        }
                                    }
                                }

                                if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "101")
                                {
                                    query = "EXEC PS_PP070_04 '" + oMat01.Columns.Item("PP030No").Cells.Item(pVal.Row).Specific.Value + "'";
                                }
                                else
                                {
                                    query = "EXEC PS_PP080_02 '" + oMat01.Columns.Item("PP030No").Cells.Item(pVal.Row).Specific.Value + "','" + oForm.Items.Item("OrdGbn").Specific.Selected.Value + "'";
                                }

                                RecordSet01.DoQuery(query);

                                if (RecordSet01.RecordCount == 0)
                                {
                                    oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                }
                                else
                                {
                                    oDS_PS_PP080L.SetValue("U_PP030No", pVal.Row - 1, RecordSet01.Fields.Item("PP030No").Value);
                                    oDS_PS_PP080L.SetValue("U_OrdGbn", pVal.Row - 1, RecordSet01.Fields.Item("OrdGbn").Value);
                                    oDS_PS_PP080L.SetValue("U_OrdNum", pVal.Row - 1, RecordSet01.Fields.Item("OrdNum").Value);
                                    oDS_PS_PP080L.SetValue("U_OrdSub1", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub1").Value);
                                    oDS_PS_PP080L.SetValue("U_OrdSub2", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub2").Value);
                                    oDS_PS_PP080L.SetValue("U_PP030HNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030HNo").Value);
                                    oDS_PS_PP080L.SetValue("U_PP030MNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030MNo").Value);
                                    oDS_PS_PP080L.SetValue("U_ORDRNo", pVal.Row - 1, RecordSet01.Fields.Item("ORDRNo").Value);
                                    oDS_PS_PP080L.SetValue("U_RDR1No", pVal.Row - 1, RecordSet01.Fields.Item("RDR1No").Value);
                                    oDS_PS_PP080L.SetValue("U_BPLId", pVal.Row - 1, RecordSet01.Fields.Item("BPLId").Value);
                                    oDS_PS_PP080L.SetValue("U_ItemCode", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
                                    oDS_PS_PP080L.SetValue("U_ItemName", pVal.Row - 1, RecordSet01.Fields.Item("ItemName").Value);

                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104") //Multi일 경우만 온산 Lot 존재
                                    {
                                        oDS_PS_PP080L.SetValue("U_OnsanLot", pVal.Row - 1, RecordSet01.Fields.Item("OnsanLot").Value);
                                    }
                                    oDS_PS_PP080L.SetValue("U_CpCode", pVal.Row - 1, RecordSet01.Fields.Item("CpCode").Value);
                                    oDS_PS_PP080L.SetValue("U_CpName", pVal.Row - 1, RecordSet01.Fields.Item("CpName").Value);
                                    oDS_PS_PP080L.SetValue("U_BQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                    oDS_PS_PP080L.SetValue("U_BWeight", pVal.Row - 1, RecordSet01.Fields.Item("BWeight").Value);
                                    oDS_PS_PP080L.SetValue("U_PQty", pVal.Row - 1, RecordSet01.Fields.Item("PQty").Value);
                                    oDS_PS_PP080L.SetValue("U_PWeight", pVal.Row - 1, RecordSet01.Fields.Item("PWeight").Value);
                                    oDS_PS_PP080L.SetValue("U_YQty", pVal.Row - 1, RecordSet01.Fields.Item("YQty").Value);
                                    oDS_PS_PP080L.SetValue("U_YWeight", pVal.Row - 1, RecordSet01.Fields.Item("YWeight").Value);
                                    oDS_PS_PP080L.SetValue("U_NQty", pVal.Row - 1, RecordSet01.Fields.Item("NQty").Value);
                                    oDS_PS_PP080L.SetValue("U_NWeight", pVal.Row - 1, RecordSet01.Fields.Item("NWeight").Value);
                                    oDS_PS_PP080L.SetValue("U_WhsCode", pVal.Row - 1, RecordSet01.Fields.Item("WhsCode").Value);
                                    oDS_PS_PP080L.SetValue("U_WhsName", pVal.Row - 1, RecordSet01.Fields.Item("WhsName").Value);
                                    oDS_PS_PP080L.SetValue("U_BatchNum", pVal.Row - 1, RecordSet01.Fields.Item("BatchNum").Value);
                                    oDS_PS_PP080L.SetValue("U_LineId", pVal.Row - 1, RecordSet01.Fields.Item("LineId").Value);
                                }

                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP080_AddMatrixRow(pVal.Row, false);
                                }

                                oMat01.LoadFromDataSource();
                            }
                            else if (pVal.ColUID == "PQty")
                            {
                                oMat01.FlushToDataSource();
                                if (oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "104") //멀티
                                {
                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                    }
                                    else if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) != Convert.ToDouble(oMat01.Columns.Item("YQty").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                    }
                                    else
                                    {
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        oDS_PS_PP080L.SetValue("U_YQty", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        oDS_PS_PP080L.SetValue("U_NQty", pVal.Row - 1, "0");
                                    }
                                }
                                else //엔트베어링,휘팅,부품,기계,몰드
                                {
                                    string temp = oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1);

                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)); //생산수량
                                        oDS_PS_PP080L.SetValue("U_PWeight", pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)); //생산중량
                                        oDS_PS_PP080L.SetValue("U_YQty", pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)); //합격수량
                                        oDS_PS_PP080L.SetValue("U_YWeight", pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)); //합격중량
                                    }
                                    else
                                    {
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        oDS_PS_PP080L.SetValue("U_YQty", pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));

                                        if (oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "105" || oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "106")
                                        {
                                            oDS_PS_PP080L.SetValue("U_Cost01", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("Cost01").Specific.Value) - Convert.ToDouble(oForm.Items.Item("aCost01").Specific.Value)));
                                            oDS_PS_PP080L.SetValue("U_Cost02", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("Cost02").Specific.Value) - Convert.ToDouble(oForm.Items.Item("aCost02").Specific.Value)));
                                            oDS_PS_PP080L.SetValue("U_Cost03", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("Cost03").Specific.Value) - Convert.ToDouble(oForm.Items.Item("aCost03").Specific.Value)));
                                            oDS_PS_PP080L.SetValue("U_Cost04", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("Cost04").Specific.Value) - Convert.ToDouble(oForm.Items.Item("aCost04").Specific.Value)));
                                            oDS_PS_PP080L.SetValue("U_Cost05", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("Cost05").Specific.Value) - Convert.ToDouble(oForm.Items.Item("aCost05").Specific.Value)));
                                            oDS_PS_PP080L.SetValue("U_Cost06", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("Cost06").Specific.Value) - Convert.ToDouble(oForm.Items.Item("aCost06").Specific.Value)));
                                            oDS_PS_PP080L.SetValue("U_CostTot", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("CostTot").Specific.Value) - Convert.ToDouble(oForm.Items.Item("aCostTot").Specific.Value)));

                                            if (Convert.ToDouble(oForm.Items.Item("SjAmt").Specific.Value) == 0)
                                            {
                                                oDS_PS_PP080L.SetValue("U_SaleAmt", pVal.Row - 1, "0");
                                            }
                                            else
                                            {
                                                oDS_PS_PP080L.SetValue("U_SaleAmt", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("SjAmt").Specific.Value) / Convert.ToDouble(oForm.Items.Item("SjQty").Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            }
                                        }

                                        if (oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "101")
                                        {
                                            weight = Convert.ToDouble(dataHelpClass.GetValue("SELECT U_UnWeight  FROM [OITM] WHERE ItemCode = '" + oDS_PS_PP080L.GetValue("U_ItemCode" + pVal.ColUID, pVal.Row - 1) + "'", 0, 1)) / 1000;
                                        }
                                        else
                                        {
                                            weight = 0;
                                        }
                                        if (weight == 0)
                                        {
                                            oDS_PS_PP080L.SetValue("U_PWeight", pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                            oDS_PS_PP080L.SetValue("U_YWeight", pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        }
                                        else
                                        {
                                            oDS_PS_PP080L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1))));
                                            oDS_PS_PP080L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1))));
                                        }
                                        oDS_PS_PP080L.SetValue("U_NQty", pVal.Row - 1, "0");
                                        oDS_PS_PP080L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                    }
                                }

                                oMat01.LoadFromDataSourceEx(); //Matrix Focus 고정
                            }
                            else if (pVal.ColUID == "NQty")
                            {
                                oMat01.FlushToDataSource();
                                if (oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "104") //멀티
                                {
                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > 0)
                                    {
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                                    }
                                    else if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                                    }
                                    else
                                    {
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        oDS_PS_PP080L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    }
                                }
                                else //엔트베어링,휘팅,부품,기계,몰드 //불량수량은 합격수량에 영향 없음
                                {
                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                    }
                                    else if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP080L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                    }
                                    else
                                    {
                                        oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                        if (oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "101") //휘팅
                                        {
                                            weight = Convert.ToDouble(dataHelpClass.GetValue("SELECT U_UnWeight  FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) / 1000;
                                        }
                                        else
                                        {
                                            weight = 0;
                                        }
                                        if (weight == 0)
                                        {
                                            oDS_PS_PP080L.SetValue("U_NWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        }
                                        else
                                        {
                                            oDS_PS_PP080L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                    }
                                }
                                oMat01.LoadFromDataSourceEx(); //Matrix Focus 고정
                            }
                            else
                            {
                                oDS_PS_PP080L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP080H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "CardCode")
                            {
                                oDS_PS_PP080H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_PP080H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
                            }
                            else if (pVal.ItemUID == "CntcCode")
                            {
                                oDS_PS_PP080H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_PP080H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_PP080H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                        }

                        oMat01.AutoResizeColumns();
                        oForm.Update();
                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "PP030No" || pVal.ColUID == "PQty" || pVal.ColUID == "NQty")
                            {
                                for (i = 0; i <= oMat01.VisualRowCount - 2; i++) //빈행이 추가되었기 때문에 -2를 적용
                                {
                                    sumQty += Convert.ToDouble(oDS_PS_PP080L.GetValue("U_PQty", i)); //생산수량 SUM
                                }

                                oForm.Items.Item("SumQty").Specific.Value = Convert.ToString(sumQty);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_EVENT_MATRIX_LOAD
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            double sumQty = 0;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("PQty").Cells.Item(i + 1).Specific.Value))
                        {
                            sumQty += 0;
                        }
                        else
                        {
                            sumQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i + 1).Specific.Value);
                        }
                    }

                    oForm.Items.Item("SumQty").Specific.Value = Convert.ToString(sumQty);

                    PS_PP080_FormItemEnabled();
                    PS_PP080_AddMatrixRow(oMat01.VisualRowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP080H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP080L);
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
            SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (oDataTable01 != null)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "WhsCode")
                            {
                                oDS_PS_PP080L.SetValue("U_WhsCode", pVal.Row - 1, oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value);
                                oDS_PS_PP080L.SetValue("U_WhsName", pVal.Row - 1, oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value);
                                oMat01.LoadFromDataSource();
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                if (oDataTable01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable01);
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
            double sumQty = 0;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (PS_PP080_Validate("행삭제") == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (int i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }

                        oMat01.FlushToDataSource();
                        oDS_PS_PP080L.RemoveRecord(oDS_PS_PP080L.Size - 1);
                        oMat01.LoadFromDataSource();

                        if (oMat01.RowCount == 0)
                        {
                            PS_PP080_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_PP080L.GetValue("U_PP030No", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_PP080_AddMatrixRow(oMat01.RowCount, false);
                            }
                            
                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                sumQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i + 1).Specific.Value); //생산수량 sum
                            }

                            oForm.Items.Item("SumQty").Specific.Value = Convert.ToString(sumQty);
                        }
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
            int rowCount = 0;
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 취소할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
                                BubbleEvent = false;
                                return;
                            }
                            for (int i = 1; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                if (oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked == true && string.IsNullOrEmpty(oMat01.Columns.Item("OIGENum").Cells.Item(i).Specific.Value.ToString().Trim()))
                                {
                                    rowCount += 1;
                                }
                            }

                            if (rowCount == 0)
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("취소할 항목을 선택해주세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_PP080_Validate("취소") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", 1, "예", "아니오") != 1)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_PP080_DI_API03() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            break;
                        case "1286": //닫기
                            if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 닫기할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
                            sQry = "Select Min(IsNULL(U_OIGENum, '')) From [@PS_PP080L] where DocEntry = '" + oDocEntry + "'";
                            oRecordSet01.DoQuery(sQry);

                            if (string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value.ToString().Trim()))
                            {
                                oStatus = "O";
                                oCanceled = "N";
                            }
                            else
                            {
                                oStatus = "C";
                                oCanceled = "Y";
                            }

                            dataHelpClass.DoQuery("UPDATE [@PS_PP080H] SET Status = '" + oStatus + "', Canceled = '" + oCanceled + "' WHERE DocEntry = '" + oDocEntry + "'");

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PS_PP080_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Specific.Value = oDocEntry;
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_PP080_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP080_FormItemEnabled();
                            PS_PP080_AddMatrixRow(0, true);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_PP080_FormItemEnabled();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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

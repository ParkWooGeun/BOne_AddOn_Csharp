using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using SAP.Middleware.Connector;
using PSH_BOne_AddOn.Code;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 입고(배치번호관리품목)
	/// </summary>
	internal class PS_MM180 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01; 
		private SAPbouiCOM.DBDataSource oDS_PS_MM180H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_MM180L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        protected class ItemInformation
        {
            public string ItemCode;
            public string BatchNum;
            public string WhsCode;
            public double Quantity;
            public bool Check;
            public int OIGNNo;
            public int IGN1No;
            public int IGE1No;
            public int MM180HNum;
            public int MM180LNum;
        }

        /// <summary>
        /// LoadForm
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM180.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM180_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM180");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_MM180_CreateItems();
				PS_MM180_CF_ChooseFromList();
				PS_MM180_EnableMenus();
				PS_MM180_SetDocument(oFormDocEntry);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
				oForm.Freeze(false);
			}
		}

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_MM180_CreateItems()
        {
            try
            {
                oDS_PS_MM180H = oForm.DataSources.DBDataSources.Item("@PS_MM180H");
                oDS_PS_MM180L = oForm.DataSources.DBDataSources.Item("@PS_MM180L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("BItemCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("BItemCod").Specific.DataBind.SetBound(true, "", "BItemCod");

                oForm.DataSources.UserDataSources.Add("BWhsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 15);
                oForm.Items.Item("BWhsCode").Specific.DataBind.SetBound(true, "", "BWhsCode");

                oForm.DataSources.UserDataSources.Add("BItemNam", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("BItemNam").Specific.DataBind.SetBound(true, "", "BItemNam");

                oForm.DataSources.UserDataSources.Add("BWhsName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("BWhsName").Specific.DataBind.SetBound(true, "", "BWhsName");

                oForm.DataSources.UserDataSources.Add("BBatchNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 32);
                oForm.Items.Item("BBatchNm").Specific.DataBind.SetBound(true, "", "BBatchNm");

                oForm.DataSources.UserDataSources.Add("BoxNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 32);
                oForm.Items.Item("BoxNo").Specific.DataBind.SetBound(true, "", "BoxNo");

                oForm.DataSources.UserDataSources.Add("BBatchSt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 29);
                oForm.Items.Item("BBatchSt").Specific.DataBind.SetBound(true, "", "BBatchSt");

                oForm.DataSources.UserDataSources.Add("BBatchEd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 29);
                oForm.Items.Item("BBatchEd").Specific.DataBind.SetBound(true, "", "BBatchEd");

                oForm.DataSources.UserDataSources.Add("BQuantity", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("BQuantity").Specific.DataBind.SetBound(true, "", "BQuantity");

                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList
        /// </summary>
        private void PS_MM180_CF_ChooseFromList()
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Condition oCon = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.Column oColumn = null;

            try
            {
                oEdit = oForm.Items.Item("BWhsCode").Specific;
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLCreationParams.ObjectType = "64";
                oCFLCreationParams.UniqueID = "CFLWHSCODE";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);
                oEdit.ChooseFromListUID = "CFLWHSCODE";
                oEdit.ChooseFromListAlias = "WhsCode";

                oColumn = oMat01.Columns.Item("WhsCode");
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLCreationParams.ObjectType = "64";
                oCFLCreationParams.UniqueID = "CFLWHSCODE2";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);
                oColumn.ChooseFromListUID = "CFLWHSCODE2";
                oColumn.ChooseFromListAlias = "WhsCode";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oCFLs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                }

                if (oCons != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCons);
                }

                if (oCon != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCon);
                }
                
                if (oCFL != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                }
                
                if (oCFLCreationParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                }
                
                if (oEdit != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
                }
                
                if (oColumn != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn);
                }                
            }
        }

        /// <summary>
		/// EnableMenus
		/// </summary>
        private void PS_MM180_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, true, false, true, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면 초기화
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PS_MM180_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_MM180_EnableFormItem();
                    PS_MM180_AddMatrixRow(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_MM180_EnableFormItem();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각 모드에 따른 아이템설정
        /// </summary>
        private void PS_MM180_EnableFormItem()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Button01").Enabled = true;
                    oForm.Items.Item("Button02").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;

                    oForm.Items.Item("BItemCod").Specific.Value = "";
                    oForm.Items.Item("BWhsCode").Specific.Value = "";
                    oForm.Items.Item("BoxNo").Specific.Value = "";
                    oForm.Items.Item("BBatchNm").Specific.Value = "";
                    oForm.Items.Item("BBatchSt").Specific.Value = "";
                    oForm.Items.Item("BBatchEd").Specific.Value = "";
                    oForm.Items.Item("BQuantity").Specific.Value = 0;

                    PS_MM180_SetUserDataSourceItem();
                    PS_MM180_SetDocEntry();
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Button01").Enabled = false;
                    oForm.Items.Item("Button02").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = false;

                    PS_MM180_SetUserDataSourceItem();
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("Button01").Enabled = false;
                    oForm.Items.Item("Button02").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = false;
                }
                oMat01.AutoResizeColumns();
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_MM180_SetDocEntry()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                string DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM180'", "");
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
        /// 메트릭스 행추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_MM180_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM180L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM180L.Offset = oRow;
                oDS_PS_MM180L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
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
        /// 필수 입력 사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_MM180_CheckDataValid()
        {
            bool returnValue = false;
            string sQry;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errMessage = "전기일은 필수입니다.";
                    oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (dataHelpClass.Future_Date_Check(oForm.Items.Item("DocDate").Specific.Value) == "N")
                {
                    errMessage = "미래일자는 입력할 수 없습니다.";
                    oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 1)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }
                
                for (int i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "품목은 필수입니다.";
                        oMat01.Columns.Item("ItemCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    if (string.IsNullOrEmpty(oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "창고는 필수입니다.";
                        oMat01.Columns.Item("WhsCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    if (Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "중량은 필수입니다.";
                        oMat01.Columns.Item("Quantity").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    if (dataHelpClass.GetItem_ManBtchNum(oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value) == "Y") //배치존재유무
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "배치는 필수입니다.";
                            oMat01.Columns.Item("BatchNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }

                        if (oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value.ToString().ToUpper() != oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value)
                        {
                            errMessage = "소문자로 입력하셨습니다. 배치번호를 확인하십시오.";
                            oMat01.Columns.Item("BatchNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    
                        sQry = "  SELECT    COUNT(*) ";
                        sQry += " FROM      [OBTN] AS T0";
                        sQry += "           LEFT JOIN";
                        sQry += "           [OBTQ] AS T1";
                        sQry += "               ON T0.ItemCode = T1.ItemCode";
                        sQry += "               AND T0.SysNumber = T1.SysNumber";
                        sQry += " WHERE     T0.DistNumber = '" + oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value + "'";
                        sQry += "           AND T0.ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value + "'";
                        sQry += "           AND T1.WhsCode = '" + oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value + "'";
                        sQry += "           AND T1.Quantity > 0";

                        if (Convert.ToInt32(dataHelpClass.GetValue(sQry, 0, 1)) > 0)
                        {
                            errMessage = "해당품목의 가용배치가 존재합니다.";
                            oMat01.Columns.Item("ItemCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    
                        sQry = "  select    COUNT(*)";
                        sQry += " from      [@PS_PP040H] a";
                        sQry += "           Inner Join";
                        sQry += "           [@PS_PP040L] b";
                        sQry += "               On a.DocEntry = b.DocEntry";
                        sQry += "               And a.Canceled = 'N'";
                        sQry += "               AND a.U_DocDate BETWEEN DATEADD(YY,-9,GETDATE()) AND GETDATE() ";
                        sQry += "           Inner Join";
                        sQry += "           [@PS_PP030L] c";
                        sQry += "               On b.U_PP030HNo = c.DocEntry";
                        sQry += " Where     c.U_BatchNum = '" + oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value + "'";

                        if (Convert.ToInt32(dataHelpClass.GetValue(sQry, 0, 1)) > 0)
                        {
                            errMessage = "이미 생산에 투입된 배치번호입니다.";
                            oMat01.Columns.Item("ItemCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    
                        for (int j = i + 1; j <= oMat01.VisualRowCount - 1; j++)
                        {
                            if (oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value == oMat01.Columns.Item("ItemCode").Cells.Item(j).Specific.Value) //품목코드가 같고
                            {
                                if (oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("BatchNum").Cells.Item(j).Specific.Value) //배치번호가 같으면
                                {
                                    errMessage = "중복된 품목,배치가 존재합니다.";
                                    oMat01.Columns.Item("ItemCode").Cells.Item(j).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    throw new Exception();
                                }
                            }
                        }
                    }
                }

                if (string.IsNullOrEmpty(oMat01.Columns.Item("BatchNum").Cells.Item(oMat01.VisualRowCount).Specific.Value))
                {
                    oDS_PS_MM180L.RemoveRecord(oDS_PS_MM180L.Size - 1);
                }
                oMat01.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_MM180_SetDocEntry();
                }

                returnValue = true;
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
            }
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_MM180_Validate(string ValidateType)
        {
            bool returnValue = false;
            string sQry;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_MM180H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    errMessage = "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할 수 없습니다.";
                    throw new Exception();
                }

                if (ValidateType == "수정")
                {
                }
                else if (ValidateType == "행삭제")
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        errMessage = "해당모드에서는 행삭제를 할 수 없습니다.";
                        throw new Exception();
                    }
                }
                else if (ValidateType == "취소")
                {
                    sQry = "SELECT U_ItemCode AS ItemCode, U_BatchNum AS BatchNum, U_WhsCode AS WhsCode, U_Quantity AS Quantity FROM [@PS_MM180L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                    RecordSet01.DoQuery(sQry);
                    for (int i = 0; i <= RecordSet01.RecordCount - 1; i++)
                    {
                        if (dataHelpClass.GetItem_ManBtchNum(RecordSet01.Fields.Item(0).Value) == "Y")
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [OIBT] WHERE ItemCode = '" + RecordSet01.Fields.Item("ItemCode").Value + "' AND BatchNum = '" + RecordSet01.Fields.Item("BatchNum").Value + "' AND WhsCode = '" + RecordSet01.Fields.Item("WhsCode").Value + "' AND Quantity = " + RecordSet01.Fields.Item("Quantity").Value + "", 0, 1)) <= 0)
                            {
                                errMessage = "품목코드 : " + RecordSet01.Fields.Item(0).Value + ", 배치번호 : " + RecordSet01.Fields.Item(1).Value + " 의 재고가 존재하지 않습니다.";
                                throw new Exception();
                            }
                        }
                        RecordSet01.MoveNext();
                    }
                }

                returnValue = true;
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
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 합계 계산
        /// </summary>
        private void PS_MM180_CalculateSumQty()
        {
            double SumQty = 0;

            try
            {
                for (int i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    SumQty += Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value);
                }

                oForm.Items.Item("SumQty").Specific.Value = SumQty;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Batch 생성
        /// </summary>
        private void PS_MM180_CreateBatch()
        {
            int i;
            int ValidBatch;
            int StartValue;
            int EndValue;
            string BatchNum;
            string errMessage = string.Empty;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("BItemCod").Specific.Value))
                {
                    errMessage = "품목코드를 선택하지 않았습니다.";
                    oForm.Items.Item("BItemCod").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("BWhsCode").Specific.Value))
                {
                    errMessage = "창고코드를 선택하지 않았습니다.";
                    oForm.Items.Item("BWhsCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("BBatchNm").Specific.Value))
                {
                    errMessage = "배치코드를 선택하지 않았습니다.";
                    oForm.Items.Item("BBatchNm").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("BBatchSt").Specific.Value))
                {
                    errMessage = "배치시작번호를 선택하지 않았습니다.";
                    oForm.Items.Item("BBatchSt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("BBatchEd").Specific.Value))
                {
                    errMessage = "배치끝번호를 선택하지 않았습니다.";
                    oForm.Items.Item("BBatchEd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (Convert.ToDouble(oForm.Items.Item("BQuantity").Specific.Value) <= 0)
                {
                    errMessage = "중량을 선택하지 않았습니다.";
                    oForm.Items.Item("BQuantity").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (oForm.Items.Item("BBatchNm").Specific.Value.ToString().Length != 8)
                {
                    errMessage = "배치코드는 8자리 입니다. 확인하여 주십시오.";
                    oForm.Items.Item("BBatchNm").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                StartValue = Convert.ToInt32(codeHelpClass.Right(oForm.Items.Item("BBatchSt").Specific.Value, 3));
                EndValue = Convert.ToInt32(codeHelpClass.Right(oForm.Items.Item("BBatchEd").Specific.Value, 3));

                if (StartValue == 0 && EndValue == 0)
                {
                    errMessage = "배치번호의 범위가 올바르지 않습니다.";
                    throw new Exception();
                }

                ValidBatch = EndValue - StartValue;

                if (ValidBatch < 0)
                {
                    errMessage = "배치번호의 범위가 올바르지 않습니다.";
                    throw new Exception();
                }

                oForm.Freeze(true);

                ValidBatch = (ValidBatch / 10) + 1; //10단위로 몇개생성가능한지 계산
                int MatrixRow = oMat01.VisualRowCount;
                for (i = 1; i <= ValidBatch; i++)
                {
                    BatchNum = Convert.ToString(StartValue + (10 * (i - 1)));
                    if (Convert.ToInt32(BatchNum) < 100)
                    {
                        BatchNum = "0" + BatchNum;
                    }

                    oDS_PS_MM180L.SetValue("U_ItemCode", MatrixRow - 1, oForm.Items.Item("BItemCod").Specific.Value);
                    oDS_PS_MM180L.SetValue("U_ItemName", MatrixRow - 1, dataHelpClass.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item("BItemCod").Specific.Value + "'", 0, 1));
                    oDS_PS_MM180L.SetValue("U_BatchNum", MatrixRow - 1, (oForm.Items.Item("BBatchNm").Specific.Value + BatchNum).ToString().ToUpper());
                    oDS_PS_MM180L.SetValue("U_BoxNo", MatrixRow - 1, oForm.Items.Item("BoxNo").Specific.Value);
                    oDS_PS_MM180L.SetValue("U_WhsCode", MatrixRow - 1, oForm.Items.Item("BWhsCode").Specific.Value);
                    oDS_PS_MM180L.SetValue("U_WhsName", MatrixRow - 1, dataHelpClass.GetValue("SELECT WhsName FROM [OWHS] WHERE WhsCode = '" + oForm.Items.Item("BWhsCode").Specific.Value + "'", 0, 1));
                    oDS_PS_MM180L.SetValue("U_Quantity", MatrixRow - 1, oForm.Items.Item("BQuantity").Specific.Value);
                    PS_MM180_AddMatrixRow(MatrixRow, false);
                    MatrixRow += 1;
                }

                PS_MM180_CalculateSumQty();

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
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
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// UserDataSource Item 초기화
        /// </summary>
        private void PS_MM180_SetUserDataSourceItem()
        {
            try
            {
                oForm.DataSources.UserDataSources.Item("BItemCod").Value = "";
                oForm.DataSources.UserDataSources.Item("BWhsCode").Value = "";
                oForm.DataSources.UserDataSources.Item("BItemNam").Value = "";
                oForm.DataSources.UserDataSources.Item("BWhsName").Value = "";
                oForm.DataSources.UserDataSources.Item("BBatchNm").Value = "";
                oForm.DataSources.UserDataSources.Item("BoxNo").Value = "";
                oForm.DataSources.UserDataSources.Item("BBatchSt").Value = "";
                oForm.DataSources.UserDataSources.Item("BBatchEd").Value = "";
                oForm.DataSources.UserDataSources.Item("BQuantity").Value = "0";
                oForm.DataSources.UserDataSources.Item("SumQty").Value = "0";
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 본사 데이터 전송
        /// </summary>
        private bool PS_MM180_InterfaceB1toR3()
        {
            bool returnValue = false;
            string sQry;
            string Client; //클라이언트
            string ServerIP; //서버IP
            string errCode = string.Empty;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            RfcDestination rfcDest = null;
            RfcRepository rfcRep = null;

            try
            {
                oMat01.FlushToDataSource();

                Client = dataHelpClass.GetR3ServerInfo()[0];
                ServerIP = dataHelpClass.GetR3ServerInfo()[1];

                //0. 연결
                if (dataHelpClass.SAPConnection(ref rfcDest, ref rfcRep, "PSC", ServerIP, Client, "ifuser", "pdauser") == false)
                {
                    errCode = "1";
                    throw new Exception();
                }

                //1. SAP R3 함수 호출(매개변수 전달)
                IRfcFunction oFunction = rfcRep.CreateFunction("ZPP_HOLDINGS_INTF_GR");

                if (string.IsNullOrEmpty(oDS_PS_MM180L.GetValue("U_SPNUM", 0).ToString().Trim())) // 박스번호로 입력시
                {
                    for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        oFunction.SetValue("I_ZLOTNO", oDS_PS_MM180L.GetValue("U_BatchNum", i).ToString().Trim()); //입고로트번호
                        oFunction.SetValue("I_RSDAT", oDS_PS_MM180H.GetValue("U_DocDate", 0)); //입고일자

                        errCode = "2"; //SAP Function 실행 오류가 발생했을 때 에러코드로 처리하기 위해 이 위치에서 "2"를 대입
                        oFunction.Invoke(rfcDest); //Function 실행

                        if (oFunction.GetValue("E_MESSAGE").ToString().Trim() != "" && codeHelpClass.Left(oFunction.GetValue("E_MESSAGE").ToString().Trim(), 1) != "S") //리턴 메시지가 "S(성공)"이 아니면
                        {
                            errCode = "3";
                            errMessage = oFunction.GetValue("E_MESSAGE").ToString();
                            throw new Exception();
                        }
                        else
                        {
                            oDS_PS_MM180L.SetValue("U_TransYN", i, "Y");
                        }
                    }
                }
                else // 출하번호로 입력시
                {
                    oFunction.SetValue("I_SPNUM", oDS_PS_MM180L.GetValue("U_SPNUM", 0).ToString().Trim()); //출하번호
                    oFunction.SetValue("I_RSDAT", oDS_PS_MM180H.GetValue("U_DocDate", 0)); //입고일자

                    errCode = "2"; //SAP Function 실행 오류가 발생했을 때 에러코드로 처리하기 위해 이 위치에서 "2"를 대입
                    oFunction.Invoke(rfcDest); //Function 실행

                    if (oFunction.GetValue("E_MESSAGE").ToString().Trim() != "" && codeHelpClass.Left(oFunction.GetValue("E_MESSAGE").ToString().Trim(), 1) != "S") //리턴 메시지가 "S(성공)"이 아니면
                    {
                        errCode = "3";
                        errMessage = oFunction.GetValue("E_MESSAGE").ToString();
                        throw new Exception();
                    }
                    else
                    {
                        sQry = "Update [@PS_MM180L] set U_TransYN ='Y' where DocEntry ='" + oDS_PS_MM180H.GetValue("DocEntry", 0).ToString().Trim()  + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }
                oMat01.LoadFromDataSource();
                PSH_Globals.SBO_Application.MessageBox("R3 인터페이스 완료!");
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("풍산 SAP R3에 로그온 할 수 없습니다. 관리자에게 문의 하세요.");
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("RFC Function 호출 오류");
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 1. Box No R3(울산사업장) 전송, 2. Lot No 회신, 3. Form Matrix에 출력
        /// </summary>
        private void PS_MM180_LoadBOXNoFromR3()
        {
            string E_MESSAGE;
            string sQry;
            string I_ZLOTNO;
            string I_ZPROWE;
            string I_ZBOXNO;
            string I_ZMATNR;
            string I_ZMAKTX;
            string I_ZHOEHE;
            string I_ZBREIT;
            string I_ZLAENG;
            string I_ZPONO;
            string I_ZPOMAT;
            string I_ZPOMAX;
            string I_ZTKOSD;
            string I_ZWIDTK;
            string I_KGNET;
            string I_QM_KUNNR;
            string I_NAME1;
            string I_KUNNR;
            string I_NAME2;
            string I_RDATE;
            string errMessage = string.Empty;
            string errCode = string.Empty;
            string Client; //클라이언트(운영용:210, 테스트용:810)
            string ServerIP; //서버IP(운영용:192.1.11.3, 테스트용:192.1.11.7)
            RfcRepository rfcRep = null;
            RfcDestination rfcDest = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                oForm.Freeze(true);
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                Client = dataHelpClass.GetR3ServerInfo()[0];
                ServerIP = dataHelpClass.GetR3ServerInfo()[1];

                if (string.IsNullOrEmpty(oForm.Items.Item("BWhsCode").Specific.Value))
                {
                    errMessage = "창고코드를 선택하지 않았습니다.";
                    errCode = "3";
                    oForm.Items.Item("BWhsCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                else if (oForm.Items.Item("BoxNo").Specific.Value.ToString().Length != 10)
                {
                    errMessage = "Box No는 10자리 입니다. 확인하여 주십시오.";
                    errCode = "3";
                    oForm.Items.Item("BoxNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                oMat01.FlushToDataSource();

                //0. 연결
                if (dataHelpClass.SAPConnection(ref rfcDest, ref rfcRep, "PSC", ServerIP, Client, "ifuser", "pdauser") == false)
                {
                    errCode = "1";
                    throw new Exception();
                }

                IRfcFunction oFunction = rfcRep.CreateFunction("ZPP_HOLDINGS_INTF_BOXINFO");
                I_ZBOXNO = oForm.Items.Item("BoxNo").Specific.Value;
                oFunction.SetValue("I_ZBOXNO", I_ZBOXNO); //매개변수를 문자열 변수에 저장해서 전달해야함(필수)

                int MatrixRow = 0;

                errCode = "2"; //SAP Function 실행 오류가 발생했을 때 에러코드로 처리하기 위해 이 위치에서 "2"를 대입
                oFunction.Invoke(rfcDest); //Function 실행

                IRfcTable oTable = oFunction.GetTable("ITAB");

                E_MESSAGE = oFunction.GetValue("E_MESSAGE").ToString();
                
                if (string.IsNullOrEmpty(E_MESSAGE)) //에러메시지가 없으면
                {
                    foreach (IRfcStructure row in oTable)
                    {
                        MatrixRow = oMat01.VisualRowCount;

                        I_ZLOTNO = row.GetValue("ZLOTNO").ToString();
                        I_ZPROWE = row.GetValue("ZPROWE").ToString();
                        I_ZBOXNO = row.GetValue("CHARG").ToString();
                        I_ZMATNR = row.GetValue("MATNR").ToString();
                        I_ZMAKTX = row.GetValue("MAKTX").ToString();
                        I_ZHOEHE = row.GetValue("ZHOEHE").ToString();
                        I_ZBREIT = row.GetValue("ZBREIT").ToString();
                        I_ZPONO = row.GetValue("PONO").ToString();
                        I_ZPOMAT = row.GetValue("ZPOMAT").ToString();
                        I_ZPOMAX = row.GetValue("ZPOMAX").ToString();
                        I_ZTKOSD = row.GetValue("ZTKOSD").ToString();
                        I_ZWIDTK = row.GetValue("ZWIDTK").ToString();
                        I_ZLAENG = row.GetValue("ZLAENG").ToString();
                        I_KGNET = row.GetValue("KGNET").ToString();
                        I_QM_KUNNR = row.GetValue("QM_KUNNR").ToString();
                        I_NAME1 = row.GetValue("NAME1").ToString();
                        I_KUNNR = row.GetValue("KUNNR").ToString();
                        I_NAME2 = row.GetValue("NAME2").ToString();
                        I_RDATE = row.GetValue("RDATE").ToString();

                        sQry = "Select b.U_ItemCode from [@PS_PP011H] a inner join [@PS_PP011L] b on a.Code = b.Code and a.Code ='12522' where b.U_PsCode ='" + I_ZMATNR + "'";
                        oRecordSet01.DoQuery(sQry);

                        if (oRecordSet01.Fields.Count == 0)
                        {
                            errMessage = I_ZMATNR + "원소재를 11.거래처제품(연결)코드등록에 입력하세요";
                            throw new Exception();
                        }

                        oDS_PS_MM180L.SetValue("U_ItemCode", MatrixRow - 1, oRecordSet01.Fields.Item(0).Value);
                        oDS_PS_MM180L.SetValue("U_ItemName", MatrixRow - 1, dataHelpClass.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oRecordSet01.Fields.Item(0).Value + "'", 0, 1));
                        oDS_PS_MM180L.SetValue("U_BoxNo", MatrixRow - 1, I_ZBOXNO);
                        oDS_PS_MM180L.SetValue("U_BatchNum", MatrixRow - 1, I_ZLOTNO);
                        oDS_PS_MM180L.SetValue("U_WhsCode", MatrixRow - 1, oForm.Items.Item("BWhsCode").Specific.Value);
                        oDS_PS_MM180L.SetValue("U_WhsName", MatrixRow - 1, dataHelpClass.GetValue("SELECT WhsName FROM [OWHS] WHERE WhsCode = '" + oForm.Items.Item("BWhsCode").Specific.Value + "'", 0, 1));
                        oDS_PS_MM180L.SetValue("U_Quantity", MatrixRow - 1, I_ZPROWE);

                        oDS_PS_MM180L.SetValue("U_RItemCod", MatrixRow - 1, I_ZMATNR); // R3 모재 코드
                        oDS_PS_MM180L.SetValue("U_RItemNam", MatrixRow - 1, I_ZMAKTX); // R3 모재 코드
                        oDS_PS_MM180L.SetValue("U_RThink", MatrixRow - 1, I_ZHOEHE); // R3 모재 두께L
                        oDS_PS_MM180L.SetValue("U_RWidth", MatrixRow - 1, I_ZBREIT); // R3 모재 폭
                        oDS_PS_MM180L.SetValue("U_RLength", MatrixRow - 1, I_ZLAENG); // R3 모재 길이

                        oDS_PS_MM180L.SetValue("U_R3PONum", MatrixRow - 1, I_ZPONO); // R3 완재 PO

                        oDS_PS_MM180L.SetValue("U_PItemCod", MatrixRow - 1, I_ZPOMAT); // R3 완재 코드
                        oDS_PS_MM180L.SetValue("U_PItemNam", MatrixRow - 1, I_ZPOMAX); // R3 모재 코드
                        oDS_PS_MM180L.SetValue("U_PThink", MatrixRow - 1, I_ZTKOSD); // R3 완재 두께
                        oDS_PS_MM180L.SetValue("U_PWidth", MatrixRow - 1, I_ZWIDTK); // R3 완재 폭
                        oDS_PS_MM180L.SetValue("U_PLength", MatrixRow - 1, I_ZLAENG); // R3 완재 길이

                        oDS_PS_MM180L.SetValue("U_POWeight", MatrixRow - 1, I_KGNET); // R3 PO 중량
                        oDS_PS_MM180L.SetValue("U_DestCode", MatrixRow - 1, I_QM_KUNNR); // 수주사양 거래선코드
                        oDS_PS_MM180L.SetValue("U_DestName", MatrixRow - 1, I_NAME1); // 수주사양 거래선명
                        oDS_PS_MM180L.SetValue("U_CardCode", MatrixRow - 1, I_KUNNR); // 거래처코드
                        oDS_PS_MM180L.SetValue("U_CardName", MatrixRow - 1, I_NAME2); // 거래처명

                        oDS_PS_MM180L.SetValue("U_ReqDate", MatrixRow - 1, Convert.ToDateTime(I_RDATE).ToString("yyyyMMdd")); // 요청일

                        PS_MM180_AddMatrixRow(MatrixRow, false);
                        MatrixRow += 1;
                    }
                }
                else
                {
                    errCode = "3";
                    errMessage = E_MESSAGE;
                    throw new Exception();
                }
                PS_MM180_CalculateSumQty();
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("풍산 SAP R3에 로그온 할 수 없습니다. 관리자에게 문의 하세요.");
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("RFC Function 호출 오류");
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
            }
        }


        /// <summary>
        /// 1. Box No R3(울산사업장) 전송, 2. Lot No 회신, 3. Form Matrix에 출력
        /// </summary>
        private void PS_MM180_LoadSPNUMFromR3()
        {
            string E_MESSAGE;
            string sQry;
            string I_ZSPNUM; 
            string I_ZLOTNO;
            string I_ZPROWE;
            string I_ZBOXNO;
            string I_ZMATNR;
            string I_ZMAKTX;
            string I_ZHOEHE;
            string I_ZBREIT;
            string I_ZLAENG;
            string I_ZPONO;
            string I_ZPOMAT;
            string I_ZPOMAX;
            string I_ZTKOSD;
            string I_ZWIDTK;
            string I_KGNET;
            string I_QM_KUNNR;
            string I_NAME1;
            string I_KUNNR;
            string I_NAME2;
            string I_RDATE;
            string errMessage = string.Empty;
            string errCode = string.Empty;
            string Client; //클라이언트(운영용:210, 테스트용:810)
            string ServerIP; //서버IP(운영용:192.1.11.3, 테스트용:192.1.11.7)
            RfcRepository rfcRep = null;
            RfcDestination rfcDest = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                Client = dataHelpClass.GetR3ServerInfo()[0];
                ServerIP = dataHelpClass.GetR3ServerInfo()[1];

                //if (string.IsNullOrEmpty(oForm.Items.Item("BItemCod").Specific.Value))
                //{
                //    errMessage = "품목코드를 선택하지 않았습니다.";
                //    errCode = "3";
                //    oForm.Items.Item("BItemCod").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                //    throw new Exception();
                //}

                if (string.IsNullOrEmpty(oForm.Items.Item("BWhsCode").Specific.Value))
                {
                    errMessage = "창고코드를 선택하지 않았습니다.";
                    errCode = "3";
                    oForm.Items.Item("BWhsCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                else if (oForm.Items.Item("SPNUM").Specific.Value.ToString().Length != 10)
                {
                    errMessage = "출하 계획번호는 10자리 입니다. 확인하여 주십시오.";
                    errCode = "3";
                    oForm.Items.Item("SPNUM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                oMat01.FlushToDataSource();

                //0. 연결
                if (dataHelpClass.SAPConnection(ref rfcDest, ref rfcRep, "PSC", ServerIP, Client, "ifuser", "pdauser") == false)
                {
                    errCode = "1";
                    throw new Exception();
                }

                IRfcFunction oFunction = rfcRep.CreateFunction("ZPP_HOLDINGS_INTF_BOXINFO");
                I_ZSPNUM = oForm.Items.Item("SPNUM").Specific.Value;
                oFunction.SetValue("I_SPNUM", I_ZSPNUM); //매개변수를 문자열 변수에 저장해서 전달해야함(필수)

                int MatrixRow = 0;

                errCode = "2"; //SAP Function 실행 오류가 발생했을 때 에러코드로 처리하기 위해 이 위치에서 "2"를 대입
                oFunction.Invoke(rfcDest); //Function 실행

                IRfcTable oTable = oFunction.GetTable("ITAB");

                E_MESSAGE = oFunction.GetValue("E_MESSAGE").ToString();

                if (string.IsNullOrEmpty(E_MESSAGE)) //에러메시지가 없으면
                {
                    foreach (IRfcStructure row in oTable)
                    {
                        MatrixRow = oMat01.VisualRowCount;

                        I_ZLOTNO = row.GetValue("ZLOTNO").ToString();
                        I_ZPROWE = row.GetValue("ZPROWE").ToString();
                        I_ZBOXNO = row.GetValue("CHARG").ToString();
                        I_ZMATNR = row.GetValue("MATNR").ToString();
                        I_ZMAKTX = row.GetValue("MAKTX").ToString();
                        I_ZHOEHE = row.GetValue("ZHOEHE").ToString();
                        I_ZBREIT = row.GetValue("ZBREIT").ToString();
                        I_ZPONO = row.GetValue("PONO").ToString();
                        I_ZPOMAT = row.GetValue("ZPOMAT").ToString();
                        I_ZPOMAX = row.GetValue("ZPOMAX").ToString();
                        I_ZTKOSD = row.GetValue("ZTKOSD").ToString();
                        I_ZWIDTK = row.GetValue("ZWIDTK").ToString();
                        I_ZLAENG = row.GetValue("ZLAENG").ToString();
                        I_KGNET = row.GetValue("KGNET").ToString();
                        I_QM_KUNNR = row.GetValue("QM_KUNNR").ToString();
                        I_NAME1 = row.GetValue("NAME1").ToString();
                        I_KUNNR = row.GetValue("KUNNR").ToString();
                        I_NAME2 = row.GetValue("NAME2").ToString();
                        I_RDATE = row.GetValue("RDATE").ToString();

                        sQry = "Select b.U_ItemCode from [@PS_PP011H] a inner join [@PS_PP011L] b on a.Code = b.Code and a.Code ='12522' where b.U_PsCode ='" + I_ZMATNR + "'";
                        oRecordSet01.DoQuery(sQry);

                        if (oRecordSet01.Fields.Count == 0)
                        {
                            errMessage = I_ZMATNR + "원소재를 11.거래처제품(연결)코드등록에 입력하세요";
                            throw new Exception();
                        }

                        oDS_PS_MM180L.SetValue("U_ItemCode", MatrixRow - 1, oRecordSet01.Fields.Item(0).Value);
                        oDS_PS_MM180L.SetValue("U_ItemName", MatrixRow - 1, dataHelpClass.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oRecordSet01.Fields.Item(0).Value + "'", 0, 1));
                        oDS_PS_MM180L.SetValue("U_SPNUM", MatrixRow - 1, I_ZSPNUM);
                        oDS_PS_MM180L.SetValue("U_BoxNo", MatrixRow - 1, I_ZBOXNO);
                        oDS_PS_MM180L.SetValue("U_BatchNum", MatrixRow - 1, I_ZLOTNO);
                        oDS_PS_MM180L.SetValue("U_WhsCode", MatrixRow - 1, oForm.Items.Item("BWhsCode").Specific.Value);
                        oDS_PS_MM180L.SetValue("U_WhsName", MatrixRow - 1, dataHelpClass.GetValue("SELECT WhsName FROM [OWHS] WHERE WhsCode = '" + oForm.Items.Item("BWhsCode").Specific.Value + "'", 0, 1));
                        oDS_PS_MM180L.SetValue("U_Quantity", MatrixRow - 1, I_ZPROWE);

                        oDS_PS_MM180L.SetValue("U_RItemCod", MatrixRow - 1, I_ZMATNR); // R3 모재 코드
                        oDS_PS_MM180L.SetValue("U_RItemNam", MatrixRow - 1, I_ZMAKTX); // R3 모재 코드
                        oDS_PS_MM180L.SetValue("U_RThink", MatrixRow - 1, I_ZHOEHE); // R3 모재 두께
                        oDS_PS_MM180L.SetValue("U_RWidth", MatrixRow - 1, I_ZBREIT); // R3 모재 폭
                        oDS_PS_MM180L.SetValue("U_RLength", MatrixRow - 1, I_ZLAENG); // R3 모재 길이

                        oDS_PS_MM180L.SetValue("U_R3PONum", MatrixRow - 1, I_ZPONO); // R3 완재 PO

                        oDS_PS_MM180L.SetValue("U_PItemCod", MatrixRow - 1, I_ZPOMAT); // R3 완재 코드
                        oDS_PS_MM180L.SetValue("U_PItemNam", MatrixRow - 1, I_ZPOMAX); // R3 모재 코드
                        oDS_PS_MM180L.SetValue("U_PThink", MatrixRow - 1, I_ZTKOSD); // R3 완재 두께
                        oDS_PS_MM180L.SetValue("U_PWidth", MatrixRow - 1, I_ZWIDTK); // R3 완재 폭
                        oDS_PS_MM180L.SetValue("U_PLength", MatrixRow - 1, I_ZLAENG); // R3 완재 길이

                        oDS_PS_MM180L.SetValue("U_POWeight", MatrixRow - 1, I_KGNET); // R3 PO 중량
                        oDS_PS_MM180L.SetValue("U_DestCode", MatrixRow - 1, I_QM_KUNNR); // 수주사양 거래선코드
                        oDS_PS_MM180L.SetValue("U_DestName", MatrixRow - 1, I_NAME1); // 수주사양 거래선명
                        oDS_PS_MM180L.SetValue("U_CardCode", MatrixRow - 1, I_KUNNR); // 거래처코드
                        oDS_PS_MM180L.SetValue("U_CardName", MatrixRow - 1, I_NAME2); // 거래처명

                        oDS_PS_MM180L.SetValue("U_ReqDate", MatrixRow - 1, Convert.ToDateTime(I_RDATE).ToString("yyyyMMdd")); // 요청일

                        PS_MM180_AddMatrixRow(MatrixRow, false);
                        MatrixRow += 1;
                    }
                }
                else
                {
                    errCode = "3";
                    errMessage = E_MESSAGE;
                    throw new Exception();
                }

                PS_MM180_CalculateSumQty();
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("풍산 SAP R3에 로그온 할 수 없습니다. 관리자에게 문의 하세요.");
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("RFC Function 호출 오류");
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
            }
        }

        /// <summary>
        /// 입고DI
        /// </summary>
        /// <returns></returns>
        private bool PS_MM180_DI_API_01()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int i;
            int RetVal;
            int LineNumCount;
            string MainItemCode;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents oDIObject = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;

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

                List<ItemInformation> itemInfoList = new List<ItemInformation>(); //품목정보

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    ItemInformation itemInfo = new ItemInformation
                    {
                        ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value,
                        BatchNum = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value,
                        WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value,
                        Quantity = Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value),
                        Check = false
                    };

                    itemInfoList.Add(itemInfo);
                }

                MainItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(1).Specific.Value;

                LineNumCount = 0;

                oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    oDIObject.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                }

                oDIObject.UserFields.Fields.Item("Comments").Value = "PS_MM180: 문서번호 " + oForm.Items.Item("DocEntry").Specific.Value + " 입고";

                for (i = 0; i < itemInfoList.Count; i++)
                {
                    if (i != 0)
                    {
                        oDIObject.Lines.Add();
                    }

                    oDIObject.Lines.ItemCode = itemInfoList[i].ItemCode;
                    oDIObject.Lines.WarehouseCode = itemInfoList[i].WhsCode;
                    oDIObject.Lines.Quantity = itemInfoList[i].Quantity;
                    if (dataHelpClass.GetItem_ManBtchNum(itemInfoList[i].ItemCode) == "Y")
                    {
                        oDIObject.Lines.BatchNumbers.BatchNumber = itemInfoList[i].BatchNum;
                        oDIObject.Lines.BatchNumbers.Quantity = itemInfoList[i].Quantity;
                        oDIObject.Lines.BatchNumbers.Add();
                    }
                    itemInfoList[i].IGN1No = LineNumCount;
                    LineNumCount += 1;
                }

                RetVal = oDIObject.Add();

                if (RetVal == 0)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);

                    for (i = 0; i < itemInfoList.Count; i++)
                    {
                        oDS_PS_MM180L.SetValue("U_OIGNNum", i, Convert.ToString(afterDIDocNum));
                        oDS_PS_MM180L.SetValue("U_IGN1Num", i, Convert.ToString(itemInfoList[i].IGN1No));
                    }
                }
                else
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errCode = "1";
                    throw new Exception();
                }

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

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
                else if (errCode == "3")
                {
                    //PS_MM180_InterfaceB1toR3에서 오류 발생하면 해당 메소드에서 오류 메시지 출력, 이 분기문에서는 별도 메시지 출력 안함
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

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 출고DI
        /// </summary>
        /// <returns></returns>
        private bool PS_MM180_DI_API_02()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int i;
            int RetVal;
            int LineNumCount;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents oDIObject = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;

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

                List<ItemInformation> itemInfoList = new List<ItemInformation>(); //품목정보

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    ItemInformation itemInfo = new ItemInformation
                    {
                        ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value,
                        BatchNum = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value,
                        WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value,
                        Quantity = Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value),
                        OIGNNo = Convert.ToInt32(oMat01.Columns.Item("OIGNNum").Cells.Item(i).Specific.Value),
                        MM180HNum = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value),
                        MM180LNum = Convert.ToInt32(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value),
                        Check = false
                    };

                    itemInfoList.Add(itemInfo);
                }

                LineNumCount = 0;

                oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    oDIObject.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                }
                oDIObject.UserFields.Fields.Item("Comments").Value = "PS_MM180: 문서번호 " + oForm.Items.Item("DocEntry").Specific.Value + " 입고취소";
                oDIObject.UserFields.Fields.Item("U_CancDoc").Value = itemInfoList[i - 2].OIGNNo;
                oDIObject.UserFields.Fields.Item("U_CtrlType").Value = "C";

                for (i = 0; i < itemInfoList.Count; i++)
                {
                    if (i != 0)
                    {
                        oDIObject.Lines.Add();
                    }
                    oDIObject.Lines.ItemCode = itemInfoList[i].ItemCode;
                    oDIObject.Lines.WarehouseCode = itemInfoList[i].WhsCode;
                    oDIObject.Lines.Quantity = itemInfoList[i].Quantity;

                    if (dataHelpClass.GetItem_ManBtchNum(itemInfoList[i].ItemCode) == "Y")
                    {
                        oDIObject.Lines.BatchNumbers.BatchNumber = itemInfoList[i].BatchNum;
                        oDIObject.Lines.BatchNumbers.Quantity = itemInfoList[i].Quantity;
                        oDIObject.Lines.BatchNumbers.Add();
                    }
                    itemInfoList[i].IGE1No = LineNumCount;
                    LineNumCount += 1;
                }

                RetVal = oDIObject.Add();

                if (RetVal == 0)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);

                    for (i = 0; i < itemInfoList.Count; i++)
                    {
                        dataHelpClass.DoQuery("UPDATE [@PS_MM180L] SET U_OIGENum = '" + afterDIDocNum + "', U_IGE1Num = '" + itemInfoList[i].IGE1No + "' WHERE DocEntry = '" + itemInfoList[i].MM180HNum + "' AND LineId = '" + itemInfoList[i].MM180LNum + "'");
                    }
                }
                else
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errCode = "1";
                    throw new Exception();
                }

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

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
                if (oDIObject != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
                }

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
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
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_MM180_CreateBatch();
                        }
                    }
                    if (pVal.ItemUID == "Button04")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            PS_MM180_InterfaceB1toR3();
                        }
                    }
                    else if (pVal.ItemUID == "Button03")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("SPNUM").Specific.Value) && string.IsNullOrEmpty(oForm.Items.Item("BoxNo").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.MessageBox("출하계획번호, 박스번호 중에 하나는 필수 입니다.");
                            }
                            else if (!string.IsNullOrEmpty(oForm.Items.Item("SPNUM").Specific.Value) && !string.IsNullOrEmpty(oForm.Items.Item("BoxNo").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.MessageBox("출하계획번호, 박스번호 둘다 입력할 수 없습니다.");
                            }
                            else if (!string.IsNullOrEmpty(oForm.Items.Item("BoxNo").Specific.Value))
                            {
                                PS_MM180_LoadBOXNoFromR3();
                                oMat01.Columns.Item("Quantity").Editable = false; //원소재 R3 인터페이스시 중량 수정을 막기위해 배치 로드후 중량 수정불가
                            }

                            else if (!string.IsNullOrEmpty(oForm.Items.Item("SPNUM").Specific.Value))
                            {
                                PS_MM180_LoadSPNUMFromR3();
                                oMat01.Columns.Item("Quantity").Editable = false; //원소재 R3 인터페이스시 중량 수정을 막기위해 배치 로드후 중량 수정불가
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Button02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            oForm.Freeze(true);
                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_MM180_AddMatrixRow(0, true);
                            
                            oForm.Items.Item("SumQty").Specific.Value = 0;
                            oForm.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_MM180_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                if (PS_MM180_DI_API_01() == false) //DI API 완료 후 진행
                                {
                                    BubbleEvent = false;
                                    return;
                                }

                                oMat01.LoadFromDataSource();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM180_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
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
                                PS_MM180_EnableFormItem();
                                PS_MM180_AddMatrixRow(0, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_MM180_EnableFormItem();
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
                oForm.Freeze(false);
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
                    if (pVal.ItemUID == "BItemCod")
                    {
                        if (pVal.CharPressed == 9)
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                            {
                                PS_SM010 tempForm = new PS_SM010();
                                tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                BubbleEvent = false;
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "ItemCode")
                        {
                            if (pVal.CharPressed == 9)
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PS_SM010 tempForm = new PS_SM010();
                                    tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                    BubbleEvent = false;
                                }
                            }
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "ItemCode")
                            {
                                oDS_PS_MM180L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_MM180L.SetValue("U_ItemName", pVal.Row - 1, dataHelpClass.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_MM180L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_MM180_AddMatrixRow(pVal.Row, false);
                                }
                            }
                            else if (pVal.ColUID == "Quantity")
                            {
                                PS_MM180_CalculateSumQty();
                                oDS_PS_MM180L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                            else
                            {
                                oDS_PS_MM180L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_MM180H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "BItemCod")
                            {
                                oForm.Items.Item("BItemNam").Specific.Value = dataHelpClass.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1);
                            }
                            else
                            {
                                if (pVal.ItemUID != "BWhsCode" && pVal.ItemUID != "BBatchNm" && pVal.ItemUID != "BBatchSt" && pVal.ItemUID != "BBatchEd" && pVal.ItemUID != "BQuantity" && pVal.ItemUID != "BoxNo" && pVal.ItemUID != "SPNUM") //UserDataSource Item일 경우(UDO Item이 아닐 경우)
                                {
                                    oDS_PS_MM180H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                }    
                            }
                            oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }

                        oMat01.LoadFromDataSource();
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
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
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
                    PS_MM180_SetUserDataSourceItem();
                    PS_MM180_CalculateSumQty();
                    PS_MM180_EnableFormItem();
                    PS_MM180_AddMatrixRow(oMat01.VisualRowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM180H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM180L);
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
                    if (pVal.ItemUID == "BItemCod")
                    {
                        oForm.DataSources.UserDataSources.Item("BItemCod").Value = oDataTable01.Columns.Item("ItemCode").Cells.Item(0).Value;
                        oForm.DataSources.UserDataSources.Item("BItemNam").Value = oDataTable01.Columns.Item("ItemName").Cells.Item(0).Value;
                    }
                    else if (pVal.ItemUID == "BWhsCode")
                    {
                        oForm.DataSources.UserDataSources.Item("BWhsCode").Value = oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value;
                        oForm.DataSources.UserDataSources.Item("BWhsName").Value = oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value;
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "ItemCode")
                        {
                            if (oDataTable01 != null)
                            {
                                oDS_PS_MM180L.SetValue("U_ItemCode", pVal.Row - 1, oDataTable01.Columns.Item("ItemCode").Cells.Item(0).Value);
                                oDS_PS_MM180L.SetValue("U_ItemName", pVal.Row - 1, oDataTable01.Columns.Item("ItemName").Cells.Item(0).Value);
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_MM180L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_MM180_AddMatrixRow(pVal.Row, false);
                                }
                                
                            }
                        }
                        else if (pVal.ColUID == "WhsCode")
                        {
                            if (oDataTable01 != null)
                            {
                                oDS_PS_MM180L.SetValue("U_WhsCode", pVal.Row - 1, oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value);
                                oDS_PS_MM180L.SetValue("U_WhsName", pVal.Row - 1, oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value);
                            }
                        }

                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (PS_MM180_Validate("행삭제") == false)
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
                        oDS_PS_MM180L.RemoveRecord(oDS_PS_MM180L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_MM180_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_MM180L.GetValue("U_ItemCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_MM180_AddMatrixRow(oMat01.RowCount, false);
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
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (PS_MM180_Validate("취소") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (PS_MM180_DI_API_02() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("현재 모드에서는 취소할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1286": //닫기
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
                        case "1287": //복제
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                //복제가능
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("현재 모드에서는 복제할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                BubbleEvent = false;
                                return;
                            }
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
                            PS_MM180_CalculateSumQty();
                            break;
                        case "1281": //찾기
                            PS_MM180_EnableFormItem();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_MM180_EnableFormItem();
                            PS_MM180_AddMatrixRow(0, true);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_MM180_EnableFormItem();
                            break;
                        case "1287": //복제
                            PS_MM180_EnableFormItem();
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                PS_MM180_SetDocEntry();
                            }
                            for (int i = 1; i <= oMat01.VisualRowCount; i++)
                            {
                                oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value = "";
                            }

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

using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using SAP.Middleware.Connector;
using PSH_BOne_AddOn.Code;

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

        //////사용자구조체
        //private struct ItemInformations
        //{
        //	public string ItemCode;
        //	public string BatchNum;
        //	public string WhsCode;
        //	public double Quantity;
        //	public bool Check;
        //	public int OIGNNo;
        //	public int IGN1No;
        //	public int OIGENo;
        //	public int IGE1No;
        //	public int MM180HNum;
        //	public int MM180LNum;
        //}
        //private ItemInformations[] ItemInformation;
        //private int ItemInformationCount;

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

                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_QUANTITY, 100);
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
                oCFLCreationParams.ObjectType = Convert.ToString(SAPbouiCOM.BoLinkedObject.lf_Warehouses);
                oCFLCreationParams.UniqueID = "CFLWHSCODE";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);
                oEdit.ChooseFromListUID = "CFLWHSCODE";
                oEdit.ChooseFromListAlias = "WhsCode";

                oColumn = oMat01.Columns.Item("WhsCode");
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLCreationParams.ObjectType = Convert.ToString(SAPbouiCOM.BoLinkedObject.lf_Warehouses);
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
        /// 각모드에따른 아이템설정
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
                            errMessage = "이미 생산에 투입된 배치번호 입니다.";
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
                oDS_PS_MM180L.RemoveRecord(oDS_PS_MM180L.Size - 1);
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
                    sQry = "SELECT U_ItemCode AS ItemCode,U_BatchNum AS BatchNum,U_WhsCode AS WhsCode,U_Quantity AS Quantity FROM [@PS_MM180L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
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
                for (int i = 1; i <= oMat01.VisualRowCount - 1; i++)
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
                    //else
                    //{
                    //    BatchNum = BatchNum;
                    //}
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
        /// 본사 데이터 전송
        /// </summary>
        private void PS_MM180_InterfaceB1toR3()
        {
            string Client; //클라이언트
            string ServerIP; //서버IP
            string errCode = string.Empty;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
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

                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    oFunction.SetValue("I_ZLOTNO", oDS_PS_MM180L.GetValue("U_BatchNum", i).ToString().Trim()); //입고로트번호
                    oFunction.SetValue("I_RSDAT", oDS_PS_MM180H.GetValue("U_DocDate", 0)); //입고일자

                    errCode = "2"; //SAP Function 실행 오류가 발생했을 때 에러코드로 처리하기 위해 이 위치에서 "2"를 대입
                    oFunction.Invoke(rfcDest); //Function 실행

                    if (oFunction.GetValue("E_MESSAGE").ToString() != "S") //리턴 메시지가 "S(성공)"이 아니면
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
        }

        /// <summary>
        /// 1. Box No R3(울산사업장) 전송, 2. Lot No 회신, 3. Form Matrix에 출력
        /// </summary>
        private void PS_MM180_LoadBatchFromR3()
        {
            string E_MESSAGE;
            string I_ZLOTNO;
            string I_ZPROWE;
            string errMessage = string.Empty;
            string errCode = string.Empty;
            string Client; //클라이언트(운영용:210, 테스트용:710)
            string ServerIP; //서버IP(운영용:192.1.11.3, 테스트용:192.1.11.7)

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            RfcDestination rfcDest = null;
            RfcRepository rfcRep = null;


            try
            {
                Client = dataHelpClass.GetR3ServerInfo()[0];
                ServerIP = dataHelpClass.GetR3ServerInfo()[1];

                if (string.IsNullOrEmpty(oForm.Items.Item("BItemCod").Specific.Value))
                {
                    errMessage = "품목코드를 선택하지 않았습니다.";
                    errCode = "3";
                    oForm.Items.Item("BItemCod").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("BWhsCode").Specific.Value))
                {
                    errMessage = "창고코드를 선택하지 않았습니다.";
                    errCode = "3";
                    oForm.Items.Item("BWhsCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (oForm.Items.Item("BoxNo").Specific.Value.ToString().Length != 10)
                {
                    errMessage = "배치코드는 10자리 입니다. 확인하여 주십시오.";
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
                oFunction.SetValue("I_ZBOXNO", oForm.Items.Item("BoxNo").Specific.Value);

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

                        oDS_PS_MM180L.SetValue("U_ItemCode", MatrixRow - 1, oForm.Items.Item("BItemCod").Specific.Value);
                        oDS_PS_MM180L.SetValue("U_ItemName", MatrixRow - 1, dataHelpClass.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item("BItemCod").Specific.Value + "'", 0, 1));
                        oDS_PS_MM180L.SetValue("U_BoxNo", MatrixRow - 1, oForm.Items.Item("BoxNo").Specific.Value);
                        oDS_PS_MM180L.SetValue("U_BatchNum", MatrixRow - 1, I_ZLOTNO);
                        oDS_PS_MM180L.SetValue("U_WhsCode", MatrixRow - 1, oForm.Items.Item("BWhsCode").Specific.Value);
                        oDS_PS_MM180L.SetValue("U_WhsName", MatrixRow - 1, dataHelpClass.GetValue("SELECT WhsName FROM [OWHS] WHERE WhsCode = '" + oForm.Items.Item("BWhsCode").Specific.Value + "'", 0, 1));
                        oDS_PS_MM180L.SetValue("U_Quantity", MatrixRow - 1, I_ZPROWE);

                        PS_MM180_AddMatrixRow(MatrixRow, false);
                        MatrixRow += 1;
                    }
                }
                else
                {
                    errCode = "3";
                    errMessage = E_MESSAGE;
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
            }
        }









        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	switch (pVal.EventType) {
        //		case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //			////1
        //			Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //			////2
        //			Raise_EVENT_KEY_DOWN(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //			////5
        //			Raise_EVENT_COMBO_SELECT(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CLICK:
        //			////6
        //			Raise_EVENT_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //			////7
        //			Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //			////8
        //			Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //			////10
        //			Raise_EVENT_VALIDATE(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //			////11
        //			Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pVal, ref BubbleEvent);
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
        //			Raise_EVENT_RESIZE(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //			////27
        //			Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //			////3
        //			Raise_EVENT_GOT_FOCUS(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //			////4
        //			break;
        //		////et_LOST_FOCUS
        //		case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //			////17
        //			Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //	}
        //	return;
        //	Raise_ItemEvent_Error:
        //	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        //				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					if ((PS_MM180_Validate("취소") == false)) {
        //						BubbleEvent = false;
        //						return;
        //					}
        //					if (PS_MM180_DI_API_02() == false) {
        //						BubbleEvent = false;
        //						return;
        //					}
        //				} else {
        //					MDC_Com.MDC_GF_Message(ref "현재 모드에서는 취소할수 없습니다.", ref "W");
        //					BubbleEvent = false;
        //					return;
        //				}
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
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
        //			case "1287":
        //				//복제
        //				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					////복제가능
        //				} else {
        //					MDC_Com.MDC_GF_Message(ref "현재 모드에서는 복제할수 없습니다.", ref "W");
        //					BubbleEvent = false;
        //					return;
        //				}
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((pVal.BeforeAction == false)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
        //				PS_MM180_CalculateSumQty();
        //				break;
        //			case "1281":
        //				//찾기
        //				PS_MM180_EnableFormItem();
        //				////UDO방식
        //				oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				break;
        //			case "1282":
        //				//추가
        //				PS_MM180_EnableFormItem();
        //				////UDO방식
        //				PS_MM180_AddMatrixRow(0, ref true);
        //				////UDO방식
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				PS_MM180_EnableFormItem();
        //				break;
        //			case "1287":
        //				//복제
        //				PS_MM180_EnableFormItem();
        //				if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
        //					PS_MM180_SetDocEntry();
        //				}
        //				for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value = "";
        //				}

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
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
        //		//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
        //		//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        //		//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
        //		//            MenuCreationParams01.uniqueID = "MenuUID"
        //		//            MenuCreationParams01.String = "메뉴명"
        //		//            MenuCreationParams01.Enabled = True
        //		//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
        //		//        End If
        //	} else if (pVal.BeforeAction == false) {
        //		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
        //		//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
        //		//        End If
        //	}
        //	if (pVal.ItemUID == "Mat01") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pVal.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("처리 중...", 100, false);

        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "PS_MM180") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pVal.ItemUID == "Button01") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				PS_MM180_CreateBatch();
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pVal.ItemUID == "Button03") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				PS_MM180_LoadBatchFromR3();
        //				oMat01.Columns.Item("Quantity").Editable = false;
        //				//원소재 R3 인터페이스시 중량 수정을 막기위해 배치 로드후 중량 수정불가
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pVal.ItemUID == "Button02") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				oMat01.Clear();
        //				oMat01.FlushToDataSource();
        //				oMat01.LoadFromDataSource();
        //				PS_MM180_AddMatrixRow(0, ref true);
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumQty").Specific.Value = "";
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pVal.ItemUID == "1") {

        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (PS_MM180_CheckDataValid() == false) {
        //					BubbleEvent = false;
        //					return;
        //				} else {
        //					//DI API 완료 후 진행
        //					if (PS_MM180_DI_API_01() == false) {
        //						BubbleEvent = false;
        //						return;
        //					}

        //					oMat01.LoadFromDataSource();

        //				}

        //				////해야할일 작업
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {

        //				if (PS_MM180_CheckDataValid() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				////해야할일 작업
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}

        //		}
        //	} else if (pVal.BeforeAction == false) {
        //		if (pVal.ItemUID == "PS_MM180") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pVal.ItemUID == "1") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (pVal.ActionSuccess == true) {
        //					PS_MM180_EnableFormItem();
        //					PS_MM180_AddMatrixRow(0, ref true);
        //					////UDO방식일때
        //				}
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				if (pVal.ActionSuccess == true) {
        //					PS_MM180_EnableFormItem();
        //				}
        //			}
        //		}
        //	}

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	return;
        //	Raise_EVENT_ITEM_PRESSED_Error:


        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	object ChildForm01 = null;
        //	if (pVal.BeforeAction == true) {
        //		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "ItemCode", "") '//사용자값활성
        //		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
        //		if (pVal.ItemUID == "BItemCod") {
        //			if (pVal.CharPressed == 9) {
        //				//UPGRADE_WARNING: oForm.Items(pVal.ItemUID).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value)) {
        //					ChildForm01 = new PS_SM010();
        //					//UPGRADE_WARNING: ChildForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					ChildForm01.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
        //					BubbleEvent = false;
        //					return;
        //				}
        //			}
        //		}

        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.ColUID == "ItemCode") {
        //				if (pVal.CharPressed == 9) {
        //					//UPGRADE_WARNING: oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) {
        //						ChildForm01 = new PS_SM010();
        //						//UPGRADE_WARNING: ChildForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						ChildForm01.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
        //						BubbleEvent = false;
        //						return;
        //					}
        //				}
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_KEY_DOWN_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		//        If pVal.ItemChanged = True Then
        //		//            Call oForm.Freeze(True)
        //		//            If (pVal.ItemUID = "Mat01") Then
        //		//                If (pVal.ColUID = "ItemCode") Then
        //		//                    '//기타작업
        //		//                    Call oDS_PS_MM180L.setValue("U_" & pVal.ColUID, pVal.Row - 1, oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Selected.Value)
        //		//                    If oMat01.RowCount = pVal.Row And Trim(oDS_PS_MM180L.GetValue("U_" & pVal.ColUID, pVal.Row - 1)) <> "" Then
        //		//                        PS_MM180_AddMatrixRow (pVal.Row)
        //		//                    End If
        //		//                Else
        //		//                    Call oDS_PS_MM180L.setValue("U_" & pVal.ColUID, pVal.Row - 1, oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Selected.Value)
        //		//                End If
        //		//            Else
        //		//                If (pVal.ItemUID = "CardCode") Then
        //		//                    Call oDS_PS_MM180H.setValue(pVal.ItemUID, 0, oForm.Items(pVal.ItemUID).Specific.Selected.Value)
        //		//                ElseIf (pVal.ItemUID = "CardCode") Then
        //		//                    Call oDS_PS_MM180H.setValue("U_" & pVal.ItemUID, 0, oForm.Items(pVal.ItemUID).Specific.Selected.Value)
        //		//                    Call oDS_PS_MM180H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm.Items(pVal.ItemUID).Specific.Selected.Value & "'"))
        //		//                Else
        //		//                    Call oDS_PS_MM180H.setValue("U_" & pVal.ItemUID, 0, oForm.Items(pVal.ItemUID).Specific.Selected.Value)
        //		//                End If
        //		//            End If
        //		//            oMat01.LoadFromDataSource
        //		//            oMat01.AutoResizeColumns
        //		//            oForm.Update
        //		//            If pVal.ItemUID = "Mat01" Then
        //		//                oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Click ct_Regular
        //		//            Else
        //		//                oForm.Items(pVal.ItemUID).Click ct_Regular
        //		//            End If
        //		//            Call oForm.Freeze(False)
        //		//        End If
        //	}
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_COMBO_SELECT_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CLICK
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.Row > 0) {
        //				oMat01.SelectRow(pVal.Row, true, false);
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	if (pVal.ItemUID == "Mat01") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pVal.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //	Raise_EVENT_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
        //private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_DOUBLE_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
        //private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_VALIDATE
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	short i = 0;
        //	decimal SumQty = default(decimal);
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemChanged == true) {
        //			if ((pVal.ItemUID == "Mat01")) {
        //				if ((pVal.ColUID == "ItemCode")) {
        //					////기타작업
        //					////품목코드처리
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_MM180L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_MM180L.SetValue("U_ItemName", pVal.Row - 1, MDC_PS_Common.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
        //					if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_MM180L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)))) {
        //						PS_MM180_AddMatrixRow((pVal.Row));
        //					}
        //					// 2011.01.13
        //				} else if ((pVal.ColUID == "Quantity")) {
        //					for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumQty = SumQty + oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value;
        //					}
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_MM180L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //				} else {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_MM180L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //				}
        //				oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else {
        //				if ((pVal.ItemUID == "DocEntry")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_MM180H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
        //				} else if ((pVal.ItemUID == "BItemCod")) {
        //					//UPGRADE_WARNING: oForm.Items(BItemNam).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("BItemNam").Specific.Value = MDC_PS_Common.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1);
        //				} else if ((pVal.ItemUID == "BWhsCode")) {
        //				} else if ((pVal.ItemUID == "BBatchNm")) {
        //				} else if ((pVal.ItemUID == "BBatchSt")) {
        //				} else if ((pVal.ItemUID == "BBatchEd")) {
        //				} else if ((pVal.ItemUID == "BQuantity")) {
        //				} else if ((pVal.ItemUID == "BoxNo")) {
        //				} else {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_MM180H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
        //				}
        //				oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			}
        //			oMat01.LoadFromDataSource();
        //			oMat01.AutoResizeColumns();
        //			oForm.Update();
        //			if (pVal.ItemUID == "Mat01") {
        //				oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else {
        //				oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_VALIDATE_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	short i = 0;
        //	decimal SumQty = default(decimal);
        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SumQty = SumQty + oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value;
        //		}
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //		PS_MM180_EnableFormItem();
        //		PS_MM180_AddMatrixRow(oMat01.VisualRowCount);
        //		////UDO방식
        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RESIZE
        //private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pVal = null, ref bool BubbleEvent = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		PS_MM180_ResizeForm();
        //	}
        //	return;
        //	Raise_EVENT_RESIZE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	SAPbouiCOM.DataTable oDataTable01 = null;
        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		if ((pVal.ItemUID == "BItemCod")) {
        //			//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDataTable01 = pVal.SelectedObjects;
        //			//UPGRADE_WARNING: oDataTable01.Columns().Cells().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.DataSources.UserDataSources.Item("BItemCod").Value = oDataTable01.Columns.Item("ItemCode").Cells.Item(0).Value;
        //			//UPGRADE_WARNING: oDataTable01.Columns().Cells().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.DataSources.UserDataSources.Item("BItemNam").Value = oDataTable01.Columns.Item("ItemName").Cells.Item(0).Value;
        //			//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oDataTable01 = null;
        //		}
        //		if ((pVal.ItemUID == "BWhsCode")) {
        //			//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDataTable01 = pVal.SelectedObjects;
        //			//UPGRADE_WARNING: oDataTable01.Columns().Cells().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.DataSources.UserDataSources.Item("BWhsCode").Value = oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value;
        //			//UPGRADE_WARNING: oDataTable01.Columns().Cells().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oForm.DataSources.UserDataSources.Item("BWhsName").Value = oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value;
        //			//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oDataTable01 = null;
        //		}
        //		//        If (pVal.ItemUID = "CardCode" Or pVal.ItemUID = "CardName") Then
        //		//            Call MDC_GP_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_MM180H", "U_CardCode,U_CardName")
        //		//        End If

        //		if ((pVal.ItemUID == "Mat01")) {
        //			if ((pVal.ColUID == "ItemCode")) {
        //				//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (pVal.SelectedObjects == null) {
        //				} else {
        //					//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDataTable01 = pVal.SelectedObjects;
        //					oDS_PS_MM180L.SetValue("U_ItemCode", pVal.Row - 1, oDataTable01.Columns.Item("ItemCode").Cells.Item(0).Value);
        //					oDS_PS_MM180L.SetValue("U_ItemName", pVal.Row - 1, oDataTable01.Columns.Item("ItemName").Cells.Item(0).Value);
        //					if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_MM180L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)))) {
        //						PS_MM180_AddMatrixRow(pVal.Row);
        //					}
        //					//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oDataTable01 = null;
        //					//Call MDC_GP_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
        //					oMat01.LoadFromDataSource();
        //					oMat01.AutoResizeColumns();
        //					oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				}
        //			}
        //		}
        //		if ((pVal.ItemUID == "Mat01")) {
        //			if ((pVal.ColUID == "WhsCode")) {
        //				//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (pVal.SelectedObjects == null) {
        //				} else {
        //					//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDataTable01 = pVal.SelectedObjects;
        //					oDS_PS_MM180L.SetValue("U_WhsCode", pVal.Row - 1, oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value);
        //					oDS_PS_MM180L.SetValue("U_WhsName", pVal.Row - 1, oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value);
        //					//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oDataTable01 = null;
        //					//Call MDC_GP_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
        //					oMat01.LoadFromDataSource();
        //					oMat01.AutoResizeColumns();
        //					oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				}
        //			}
        //		}
        //	}
        //	return;
        //	Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.ItemUID == "Mat01") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pVal.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //	Raise_EVENT_GOT_FOCUS_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //	} else if (pVal.BeforeAction == false) {
        //		SubMain.RemoveForms(oFormUniqueID01);
        //		//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oForm = null;
        //		//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oMat01 = null;
        //	}
        //	return;
        //	Raise_EVENT_FORM_UNLOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	if ((oLastColRow01 > 0)) {
        //		if (pVal.BeforeAction == true) {
        //			if ((PS_MM180_Validate("행삭제") == false)) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //			////행삭제전 행삭제가능여부검사
        //		} else if (pVal.BeforeAction == false) {
        //			for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //			}
        //			oMat01.FlushToDataSource();
        //			oDS_PS_MM180L.RemoveRecord(oDS_PS_MM180L.Size - 1);
        //			oMat01.LoadFromDataSource();
        //			if (oMat01.RowCount == 0) {
        //				PS_MM180_AddMatrixRow(0);
        //			} else {
        //				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_MM180L.GetValue("U_ItemCode", oMat01.RowCount - 1)))) {
        //					PS_MM180_AddMatrixRow(oMat01.RowCount);
        //				}
        //			}
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion












        #region PS_MM180_DI_API_01
        //private bool PS_MM180_DI_API_01()
        //{
        //	bool returnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	returnValue = true;
        //	object i = null;
        //	int j = 0;
        //	SAPbobsCOM.Documents oDIObject = null;
        //	int RetVal = 0;
        //	string ItmBsortCK = null;
        //	//멀티 원소재 체크 변수
        //	int LineNumCount = 0;
        //	int ResultDocNum = 0;
        //	string MainItemCode = null;
        //	//첫번째 행의 품목코드 확인변
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Company.StartTransaction();

        //	ItemInformation = new ItemInformations[1];
        //	ItemInformationCount = 0;
        //	for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //		Array.Resize(ref ItemInformation, ItemInformationCount + 1);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].BatchNum = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Quantity = oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value;
        //		ItemInformation[ItemInformationCount].Check = false;
        //		ItemInformationCount = ItemInformationCount + 1;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		MainItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(1).Specific.Value;
        //	}

        //	LineNumCount = 0;
        //	oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DocDate").Specific.Value, "&&&&-&&-&&"));
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("Comments").Value = "PS_MM180: 문서번호 " + oForm.Items.Item("DocEntry").Specific.Value + " 입고";
        //	////2010/12/21 노대리님요청
        //	for (i = 0; i <= ItemInformationCount - 1; i++) {
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (i != 0) {
        //			oDIObject.Lines.Add();
        //		}
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.ItemCode = ItemInformation[i].ItemCode;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.WarehouseCode = ItemInformation[i].WhsCode;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Quantity = ItemInformation[i].Quantity;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (MDC_PS_Common.GetItem_ManBtchNum(ItemInformation[i].ItemCode) == "Y") {
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation[i].BatchNum;
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDIObject.Lines.BatchNumbers.Quantity = ItemInformation[i].Quantity;
        //			oDIObject.Lines.BatchNumbers.Add();
        //		}
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[i].IGN1No = LineNumCount;
        //		LineNumCount = LineNumCount + 1;
        //	}
        //	RetVal = oDIObject.Add();
        //	if (RetVal == 0) {
        //		ResultDocNum = Convert.ToInt32(SubMain.Sbo_Company.GetNewObjectKey());
        //		for (i = 0; i <= Information.UBound(ItemInformation); i++) {
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDS_PS_MM180L.SetValue("U_OIGNNum", i, Convert.ToString(ResultDocNum));
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDS_PS_MM180L.SetValue("U_IGN1Num", i, Convert.ToString(ItemInformation[i].IGN1No));
        //		}
        //	} else {
        //		goto PS_MM180_DI_API_01_Error;
        //	}

        //	oMat01.LoadFromDataSource();
        //	oMat01.AutoResizeColumns();

        //	//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ItmBsortCK = MDC_GetData.Get_ReData("U_ItmBsort", "ItemCode", "[OITM]", "'" + MainItemCode + "'");



        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		//멀티 원소재 체크
        //		if (ItmBsortCK == "302") {
        //			if (PS_MM180_InterfaceB1toR3() == true) {
        //				SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //			} else {
        //				goto PS_MM180_DI_API_01_DI_Error;
        //			}


        //		//멀티 원소재 아닐경우
        //		} else {
        //			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //		}
        //	}

        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return returnValue;
        //	PS_MM180_DI_API_01_DI_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	if (PS_MM180_InterfaceB1toR3() == true) {
        //		SubMain.Sbo_Application.MessageBox("DI API 오류" + Err().Number + " - " + Err().Description + ")");
        //	}
        //	//Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
        //	returnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return returnValue;
        //	PS_MM180_DI_API_01_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_MM180_DI_API_01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	returnValue = false;
        //	return returnValue;
        //}
        #endregion

        #region PS_MM180_DI_API_02
        //private bool PS_MM180_DI_API_02()
        //{
        //	bool returnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	returnValue = true;
        //	object i = null;
        //	int j = 0;
        //	SAPbobsCOM.Documents oDIObject = null;
        //	int RetVal = 0;
        //	int LineNumCount = 0;
        //	int ResultDocNum = 0;
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Company.StartTransaction();

        //	ItemInformation = new ItemInformations[1];
        //	ItemInformationCount = 0;
        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //		Array.Resize(ref ItemInformation, ItemInformationCount + 1);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].BatchNum = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Quantity = oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].OIGNNo = oMat01.Columns.Item("OIGNNum").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].MM180HNum = Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].MM180LNum = Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value);
        //		ItemInformation[ItemInformationCount].Check = false;
        //		ItemInformationCount = ItemInformationCount + 1;
        //	}

        //	LineNumCount = 0;
        //	oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
        //	////기타출고
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DocDate").Specific.Value, "&&&&-&&-&&"));
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("Comments").Value = "PS_MM180: 문서번호 " + oForm.Items.Item("DocEntry").Specific.Value + " 입고취소";
        //	////2010/12/21 노대리님요청
        //	//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_CancDoc").Value = ItemInformation[i - 2].OIGNNo;
        //	//ItemInformation(i - 2).MM180HNum
        //	oDIObject.UserFields.Fields.Item("U_CtrlType").Value = "C";
        //	////2010/12/21 노대리님요청
        //	for (i = 0; i <= ItemInformationCount - 1; i++) {
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (i != 0) {
        //			oDIObject.Lines.Add();
        //		}
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.ItemCode = ItemInformation[i].ItemCode;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.WarehouseCode = ItemInformation[i].WhsCode;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Quantity = ItemInformation[i].Quantity;

        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (MDC_PS_Common.GetItem_ManBtchNum(ItemInformation[i].ItemCode) == "Y") {
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation[i].BatchNum;
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDIObject.Lines.BatchNumbers.Quantity = ItemInformation[i].Quantity;
        //			oDIObject.Lines.BatchNumbers.Add();
        //		}
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[i].IGE1No = LineNumCount;
        //		LineNumCount = LineNumCount + 1;
        //	}
        //	RetVal = oDIObject.Add();
        //	if (RetVal == 0) {
        //		ResultDocNum = Convert.ToInt32(SubMain.Sbo_Company.GetNewObjectKey());
        //		for (i = 0; i <= Information.UBound(ItemInformation); i++) {
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			MDC_PS_Common.DoQuery(("UPDATE [@PS_MM180L] SET U_OIGENum = '" + ResultDocNum + "', U_IGE1Num = '" + ItemInformation[i].IGE1No + "' WHERE DocEntry = '" + ItemInformation[i].MM180HNum + "' AND LineId = '" + ItemInformation[i].MM180LNum + "'"));
        //		}
        //	} else {
        //		goto PS_MM180_DI_API_02_Error;
        //	}

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //	}
        //	oMat01.LoadFromDataSource();
        //	oMat01.AutoResizeColumns();

        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return returnValue;
        //	PS_MM180_DI_API_02_DI_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage(SubMain.Sbo_Company.GetLastErrorCode() + " - " + SubMain.Sbo_Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	returnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return returnValue;
        //	PS_MM180_DI_API_02_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_MM180_DI_API_02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	returnValue = false;
        //	return returnValue;
        //}
        #endregion







    }
}

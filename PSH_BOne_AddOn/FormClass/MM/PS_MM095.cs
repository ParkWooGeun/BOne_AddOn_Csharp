using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 원자재 불출 등록
    /// </summary>
    internal class PS_MM095 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_MM095H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM095L; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_USERDS01;

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oClick_ColRow;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM095.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM095_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM095");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);
                PS_MM095_CreateItems();
                PS_MM095_ComboBox_Setting();
                PS_MM095_Initialization();
                PS_MM095_FormClear();
                PS_MM095_FormItemEnabled();

                oForm.EnableMenu(("1283"), false); // 삭제
                oForm.EnableMenu(("1286"), false); // 닫기
                oForm.EnableMenu(("1287"), false); // 복제
                oForm.EnableMenu(("1284"), true); // 취소
                oForm.EnableMenu(("1293"), true); // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
        private void PS_MM095_CreateItems()
        {
            try
            {
                oDS_PS_MM095H = oForm.DataSources.DBDataSources.Item("@PS_MM095H");
                oDS_PS_MM095L = oForm.DataSources.DBDataSources.Item("@PS_MM095L");
                oDS_PS_USERDS01 = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;

                oDS_PS_MM095H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));

                oForm.DataSources.UserDataSources.Add("Div", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Div").Specific.DataBind.SetBound(true, "", "Div");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM095_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("-", "선택");
                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where Code in ('102', '602','111') And U_PudYN = 'Y' Order by Code";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("OrdGbn").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                oForm.Items.Item("Div").Specific.ValidValues.Add("1", "전체");
                oForm.Items.Item("Div").Specific.ValidValues.Add("2", "기준일자");

                oMat01.Columns.Item("WorkGbn").ValidValues.Add("10", "자가");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("20", "정밀");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("30", "외주");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Initialization
        /// </summary>
        private void PS_MM095_Initialization()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PS_MM095H.SetValue("U_BPLId", 0, dataHelpClass.User_BPLID());
                oDS_PS_MM095H.SetValue("U_OrdGbn", 0, "-");
                oForm.DataSources.UserDataSources.Item("Div").Value = "1";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// HeaderSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM095_HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_MM095H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM095H.GetValue("U_OrdGbn", 0)))
                {
                    errMessage = "작업구분은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM095H.GetValue("U_DocDate", 0)))
                {
                    errMessage = "전기일은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return functionReturnValue;
        }

        /// <summary>
        /// MatrixSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM095_MatrixSpaceLineDel()
        {
            bool functionReturnValue = false;
            int i;
            int j = 0;
            int count;
            string errMessage = string.Empty;

            try
            {
                oForm.Freeze(true);
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }
                count = oMat01.VisualRowCount;
                for (i = 0; i <= count - 1; i++)
                {
                    oMat01.FlushToDataSource();
                    if (Convert.ToDouble(oDS_PS_MM095L.GetValue("U_IssueWt", j)) == 0)
                    {
                        oDS_PS_MM095L.RemoveRecord(j);
                        oMat01.LoadFromDataSource();
                    }
                    else
                    {
                        j++;
                    }
                }
                oMat01.LoadFromDataSource();

                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    oDS_PS_MM095L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                }
                oMat01.LoadFromDataSource();
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
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
            return functionReturnValue;
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM095_FormItemEnabled()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oMat01.Columns.Item("IssueYN").Editable = true;
                    oMat01.Columns.Item("IssueQty").Editable = true;
                    oMat01.Columns.Item("IssueWt").Editable = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oMat01.Columns.Item("IssueYN").Editable = false;
                    oMat01.Columns.Item("IssueQty").Editable = false;
                    oMat01.Columns.Item("IssueWt").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("OrdGbn").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oMat01.Columns.Item("IssueYN").Editable = false;
                    oMat01.Columns.Item("IssueQty").Editable = false;
                    oMat01.Columns.Item("IssueWt").Editable = false;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_MM095_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM095_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM095L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM095L.Offset = oRow;
                oDS_PS_MM095L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_MM095_FormClear()
        {
            string DocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM095'", "");
                if (Convert.ToDouble(DocNum) == 0)
                {
                    oForm.Items.Item("DocNum").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocNum").Specific.Value = DocNum;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// LoadData
        /// </summary>
        private void PS_MM095_LoadData()
        {
            short i;
            string sQry; 
            string BPLID;
            string OrdGbn;
            string DocDate;
            string Div;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                OrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                Div = oForm.Items.Item("Div").Specific.Value.ToString().Trim();

                sQry = "EXEC [PS_MM095_01] '" + BPLID + "','" + OrdGbn + "','" + DocDate + "','" + Div + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_MM095L.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회결과 없습니다. 확인하세요.";
                    throw new Exception();
                }
                ProgressBar01.Text = "조회시작!";
                oForm.Freeze(true);
                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_MM095L.Size)
                    {
                        oDS_PS_MM095L.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_MM095L.Offset = i;
                    oDS_PS_MM095L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_MM095L.SetValue("u_DocDate", i, oRecordSet01.Fields.Item("U_DocDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_MM095L.SetValue("U_WorkGbn", i, oRecordSet01.Fields.Item("U_WorkGbn").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_OrdNum", i, oRecordSet01.Fields.Item("U_OrdNum").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_CpCode", i, oRecordSet01.Fields.Item("U_CpCode").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_CpName", i, oRecordSet01.Fields.Item("U_CpName").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_PItemCod", i, oRecordSet01.Fields.Item("U_PItemCod").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_PItemNam", i, oRecordSet01.Fields.Item("U_PItemNam").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_PQty", i, oRecordSet01.Fields.Item("U_PQty").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_PWeight", i, oRecordSet01.Fields.Item("U_PWeight").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_CItemCod", i, oRecordSet01.Fields.Item("U_CItemCod").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_CItemNam", i, oRecordSet01.Fields.Item("U_CItemNam").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_CutCount", i, oRecordSet01.Fields.Item("U_CutCount").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_IssueYN", i, oRecordSet01.Fields.Item("U_IssueYN").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_IssueQty", i, "0");
                    oDS_PS_MM095L.SetValue("U_PreQty", i, oRecordSet01.Fields.Item("U_PreQty").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_PreWt", i, oRecordSet01.Fields.Item("U_PreWt").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_PP040Doc", i, oRecordSet01.Fields.Item("U_PP040Doc").Value.ToString().Trim());
                    oDS_PS_MM095L.SetValue("U_PP040Lin", i, oRecordSet01.Fields.Item("U_PP040Lin").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
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
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                if(ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// PS_MM095_LoadData_Mat02
        /// </summary>
        private void PS_MM095_LoadData_Mat02(int sRow)
        {
            short i;
            string sQry;
            string ItemCode;
            string CItemCod;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                ItemCode = oMat01.Columns.Item("PItemCod").Cells.Item(sRow).Specific.Value.ToString().Trim();
                CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(sRow).Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(CItemCod.ToString().Trim()))
                {
                    CItemCod = "%";
                }

                sQry = "Select a.U_ItemCod2, a.U_ItemNam2, b.WhsCode, c.WhsName, b.U_Qty, b.OnHand, ";
                sQry += " UnWeight = Case When Isnull(a.U_UnWeight,0) = 0 Then 1 When a.U_UnWeight = 0 Then 1 Else a.U_UnWeight End ";
                sQry += "From [@PS_PP005H] a Inner Join [OITW] b On a.U_ItemCod2 = b.ItemCode Inner Join [OWHS] c On b.WhsCode = c.WhsCode ";
                sQry += "Where a.U_ItemCod1 = '" + ItemCode + "' And b.OnHand > 0";
                sQry += " And a.U_ItemCod2 Like '" + CItemCod + "'";
                oRecordSet01.DoQuery(sQry);

                oMat02.Clear();
                oDS_PS_USERDS01.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회결과 없습니다. 확인하세요.";
                    throw new Exception();
                }

                oForm.Freeze(true); 
                ProgressBar01.Text = "조회시작!";

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_USERDS01.Size)
                    {
                        oDS_PS_USERDS01.InsertRecord(i);
                    }

                    oMat02.AddRow();
                    oDS_PS_USERDS01.Offset = i;
                    oDS_PS_USERDS01.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_USERDS01.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("U_ItemCod2").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColRgL01", i, oRecordSet01.Fields.Item("U_ItemNam2").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("WhsCode").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("WhsName").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColNum01", i, oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("OnHand").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColRat01", i, oRecordSet01.Fields.Item("UnWeight").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PS_MM095_Update_PP040L
        /// </summary>
        private void PS_MM095_Update_PP040L(string sStatus)
        {
            int i;
            string sQry;
            string WhsCode;
            string CItemCod;
            string IssueYN = string.Empty;
            string PP040Doc;
            string PP040Lin;
            double IssueWt;
            double IssueQty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (sStatus == "ADD")
                {
                    IssueYN = "Y";
                }
                else if (sStatus == "CANCEL")
                {
                    IssueYN = "N";
                }

                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value;
                    IssueQty = Convert.ToDouble(oMat01.Columns.Item("IssueQty").Cells.Item(i + 1).Specific.Value);
                    IssueWt = Convert.ToDouble(oMat01.Columns.Item("IssueWt").Cells.Item(i + 1).Specific.Value);
                    WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i + 1).Specific.Value;
                    if (!string.IsNullOrEmpty(CItemCod) && IssueQty >= 0 && IssueWt != 0 && !string.IsNullOrEmpty(WhsCode))
                    {
                        if (oMat01.Columns.Item("IssueYN").Cells.Item(i + 1).Specific.Value.ToString().Trim() == "Y")
                        {
                            PP040Doc = oMat01.Columns.Item("PP040Doc").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                            PP040Lin = oMat01.Columns.Item("PP040Lin").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                            sQry = "Update [@PS_PP040L] Set U_IssueYN = '" + IssueYN + "' Where DocEntry = '" + PP040Doc + "' And LineId = '" + PP040Lin + "'";
                            oRecordSet01.DoQuery(sQry);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 출고 DI
        /// </summary>
        /// <returns></returns>
        private bool PS_MM095_Add_oInventoryGenExit()
        {
            bool returnValue = false;
            int i;
            int j = 0;
            int RetVal;
            int errDICode;
            string CItemCod;
            string DocDate;
            string DocNum;
            string WhsCode;
            string errDIMsg;
            string sDocEntry;
            string errMessage = string.Empty;
            double IssueQty;
            double IssueWt;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents DI_oInventoryGenExit = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit); //문서타입(입고)
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();

                PS_MM095_FormClear();
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errMessage = "현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.";
                    throw new Exception();
                }
                DocNum = oDS_PS_MM095H.GetValue("DocNum", 0).ToString().Trim();
                DocDate = oDS_PS_MM095H.GetValue("U_DocDate", 0);

                if (string.IsNullOrEmpty(oDS_PS_MM095H.GetValue("U_OIGEDoc", 0).ToString().Trim()))
                {
                    PSH_Globals.oCompany.StartTransaction();

                    DI_oInventoryGenExit.DocDate = DateTime.ParseExact(DocDate,"yyyyMMdd",null);
                    DI_oInventoryGenExit.TaxDate = DateTime.ParseExact(DocDate, "yyyyMMdd", null);
                    DI_oInventoryGenExit.Comments = "원재료 불출 등록(" + DocNum + ") 출고 : PS_MM095";

                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value;
                        IssueQty = Convert.ToDouble(oMat01.Columns.Item("IssueQty").Cells.Item(i + 1).Specific.Value);
                        IssueWt = Convert.ToDouble(oMat01.Columns.Item("IssueWt").Cells.Item(i + 1).Specific.Value);
                        WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i + 1).Specific.Value;

                        if (!string.IsNullOrEmpty(CItemCod) && IssueQty >= 0 && IssueWt != 0 && !string.IsNullOrEmpty(WhsCode))
                        {
                            DI_oInventoryGenExit.Lines.Add();
                            DI_oInventoryGenExit.Lines.SetCurrentLine(j);
                            DI_oInventoryGenExit.Lines.ItemCode = CItemCod;
                            DI_oInventoryGenExit.Lines.WarehouseCode = WhsCode;
                            DI_oInventoryGenExit.Lines.Quantity = IssueWt;
                            DI_oInventoryGenExit.Lines.UserFields.Fields.Item("U_Qty").Value = IssueQty;
                            j++;
                        }
                    }

                    // 완료
                    RetVal = DI_oInventoryGenExit.Add();
                    if (0 != RetVal)
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
                    }
                    else
                    {
                        PSH_Globals.oCompany.GetNewObjectCode(out sDocEntry);
                        oDS_PS_MM095H.SetValue("U_OIGEDoc", 0, sDocEntry);
                        PS_MM095_Update_PP040L("ADD");
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                if (DI_oInventoryGenExit != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenExit);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// 입고 DI(취소)
        /// </summary>
        /// <returns></returns>
        private bool PS_MM095_Add_InventoryGenEntry()
        {
            bool returnValue = false;
            int i;
            int j;
            int RetVal;
            int errDICode;
            string CItemCod;
            string DocDate;
            string DocNum;
            string WhsCode;
            string errDIMsg;
            string errMessage = string.Empty;
            string sDocEntry;
            string OIGEDoc;
            double IssueQty;
            double IssueWt;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents DI_oInventoryGenEntry = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry); //문서타입(입고)
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errMessage = "현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.";
                    throw new Exception();
                }
                DocNum = oDS_PS_MM095H.GetValue("DocNum", 0).ToString().Trim();
                DocDate = oDS_PS_MM095H.GetValue("U_DocDate", 0);
                OIGEDoc = oDS_PS_MM095H.GetValue("U_OIGEDoc", 0).ToString().Trim();

                if (string.IsNullOrEmpty(oDS_PS_MM095H.GetValue("U_OIGNDoc", 0).ToString().Trim()))
                {
                    PSH_Globals.oCompany.StartTransaction();

                    DI_oInventoryGenEntry.DocDate = DateTime.ParseExact(DocDate, "yyyyMMdd", null);
                    DI_oInventoryGenEntry.TaxDate = DateTime.ParseExact(DocDate, "yyyyMMdd", null);
                    DI_oInventoryGenEntry.Comments = "원재료 불출 등록 출고 취소 (" + DocNum + ") 입고 : PS_MM095";
                    DI_oInventoryGenEntry.UserFields.Fields.Item("U_CancDoc").Value = OIGEDoc;

                    j = 0;
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value;
                        IssueQty = Convert.ToDouble(oMat01.Columns.Item("IssueQty").Cells.Item(i + 1).Specific.Value);
                        IssueWt = Convert.ToDouble(oMat01.Columns.Item("IssueWt").Cells.Item(i + 1).Specific.Value);
                        WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i + 1).Specific.Value;

                        if (!string.IsNullOrEmpty(CItemCod) && IssueQty >= 0 && IssueWt != 0 && !string.IsNullOrEmpty(WhsCode))
                        {
                            DI_oInventoryGenEntry.Lines.Add();
                            DI_oInventoryGenEntry.Lines.SetCurrentLine(j);
                            DI_oInventoryGenEntry.Lines.ItemCode = CItemCod;
                            DI_oInventoryGenEntry.Lines.WarehouseCode = WhsCode;
                            DI_oInventoryGenEntry.Lines.Quantity = IssueWt;
                            DI_oInventoryGenEntry.Lines.UserFields.Fields.Item("U_Qty").Value = IssueQty;
                            j += 1;
                        }
                    }
                    RetVal = DI_oInventoryGenEntry.Add();
                    if (0 != RetVal)
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
                    }
                    else
                    {
                        PSH_Globals.oCompany.GetNewObjectCode(out sDocEntry);
                        oDS_PS_MM095H.SetValue("U_OIGNDoc", 0, sDocEntry);
                        PS_MM095_Update_PP040L("CANCEL");
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                if (DI_oInventoryGenEntry != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenEntry);
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
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

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                //    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM095_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_MM095_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (PS_MM095_Add_oInventoryGenExit() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                PS_MM095_Update_PP040L("UPDATE");
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            //if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            //{
                            //    PS_MM095_AddMatrixRow(oMat01.RowCount, false);
                            //    oLast_Mode = 100;
                            //}
                            //else if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            //{
                            //    PS_MM095_AddMatrixRow(oMat01.RowCount, false);
                            //    PS_MM095_FormItemEnabled();
                            //    oLast_Mode = 100;
                            //}
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            oMat02.Clear();
                            oDS_PS_USERDS01.Clear();
                            oForm.Freeze(true);
                            PS_MM095_Initialization();
                            PS_MM095_FormItemEnabled();
                            PS_MM095_FormClear();
                            oForm.Items.Item("BPLId").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                            oForm.Items.Item("OrdGbn").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                            oDS_PS_MM095H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            oForm.Freeze(false);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                        if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                            if (pVal.ItemUID == "BPLId" || pVal.ItemUID == "OrdGbn" || pVal.ItemUID == "Div")
                            {
                                PS_MM095_LoadData();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "LineNum" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oClick_ColRow = pVal.Row;
                        PS_MM095_LoadData_Mat02(oClick_ColRow);
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string ItemCode;

            int Qty;
            double Weight;
            double UnWeight;
            double Calculate_Weight;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                            if (pVal.ColUID == "IssueQty")
                            {
                                oMat01.FlushToDataSource();
                                ItemCode = oDS_PS_MM095L.GetValue("U_CItemCod", pVal.Row - 1).ToString().Trim();
                                if (string.IsNullOrEmpty(ItemCode))
                                {
                                    oDS_PS_MM095L.SetValue("U_IssueQty", pVal.Row - 1, "0");
                                    oDS_PS_MM095L.SetValue("U_IssueWt", pVal.Row - 1, "0");
                                    oMat01.LoadFromDataSource();
                                    oMat01.Columns.Item("IssueQty").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    dataHelpClass.MDC_GF_Message("불출할 원자재를 아래 매트릭스에서 먼저 선택해 주세요.", "E");
                                }
                                else
                                {
                                    Qty = Convert.ToInt32(oDS_PS_MM095L.GetValue("U_CQty", pVal.Row - 1));
                                    Weight = Convert.ToDouble(oDS_PS_MM095L.GetValue("U_CWeight", pVal.Row - 1));
                                    if (Qty == 0)
                                    {
                                        UnWeight = 0;
                                    }
                                    else
                                    {
                                        UnWeight = Weight / Qty;
                                    }

                                    Calculate_Weight = Convert.ToDouble(oDS_PS_MM095L.GetValue("U_IssueQty", pVal.Row - 1)) * UnWeight;
                                    oDS_PS_MM095L.SetValue("U_IssueWt", pVal.Row - 1, Convert.ToString(Calculate_Weight));
                                    oMat01.LoadFromDataSource();
                                    oMat01.Columns.Item("IssueQty").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                            }
                            else if (pVal.ColUID == "IssueWt")
                            {
                                oMat01.FlushToDataSource();
                                ItemCode = oDS_PS_MM095L.GetValue("U_CItemCod", pVal.Row - 1).ToString().Trim();
                                if (string.IsNullOrEmpty(ItemCode))
                                {
                                    oDS_PS_MM095L.SetValue("U_IssueQty", pVal.Row - 1, "0");
                                    oDS_PS_MM095L.SetValue("U_IssueWt", pVal.Row - 1, "0");
                                    oMat01.LoadFromDataSource();
                                    oMat01.Columns.Item("IssueQty").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    dataHelpClass.MDC_GF_Message("불출할 원자재를 아래 매트릭스에서 먼저 선택해 주세요.", "E");
                                }
                                oForm.Freeze(false);

                            }
                        }
                        else if (pVal.ItemUID == "DocDate")
                        {
                            PS_MM095_LoadData();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
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
                    SubMain.Remove_Forms(oFormUniqueID);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM095H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM095L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_USERDS01);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string errMessage = string.Empty;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (!string.IsNullOrEmpty(oMat02.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value) && !string.IsNullOrEmpty(oMat02.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value) && !string.IsNullOrEmpty(oMat02.Columns.Item("WhsCode").Cells.Item(pVal.Row).Specific.Value))
                        {
                            oForm.Freeze(true);
                            oMat01.FlushToDataSource();
                            oDS_PS_MM095L.SetValue("U_CItemCod", oClick_ColRow - 1, oMat02.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value);
                            oDS_PS_MM095L.SetValue("U_CItemNam", oClick_ColRow - 1, oMat02.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value);
                            oDS_PS_MM095L.SetValue("U_WhsCode", oClick_ColRow - 1, oMat02.Columns.Item("WhsCode").Cells.Item(pVal.Row).Specific.Value);
                            oDS_PS_MM095L.SetValue("U_CQty", oClick_ColRow - 1, oMat02.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value);
                            oDS_PS_MM095L.SetValue("U_CWeight", oClick_ColRow - 1, oMat02.Columns.Item("Weight").Cells.Item(pVal.Row).Specific.Value);
                            oDS_PS_MM095L.SetValue("U_UnWeight", oClick_ColRow - 1, oMat02.Columns.Item("UnWeight").Cells.Item(pVal.Row).Specific.Value);
                            if (Convert.ToDouble(oMat02.Columns.Item("UnWeight").Cells.Item(pVal.Row).Specific.Value) == 1)
                            {
                                if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "111")
                                {
                                    oDS_PS_MM095L.SetValue("U_IssueQty", oClick_ColRow - 1, oMat01.Columns.Item("PWeight").Cells.Item(oClick_ColRow).Specific.Value);
                                    oDS_PS_MM095L.SetValue("U_IssueWt", oClick_ColRow - 1, oMat01.Columns.Item("PWeight").Cells.Item(oClick_ColRow).Specific.Value);
                                }
                                else
                                {
                                    oDS_PS_MM095L.SetValue("U_IssueQty", oClick_ColRow - 1, oMat01.Columns.Item("PQty").Cells.Item(oClick_ColRow).Specific.Value);
                                    oDS_PS_MM095L.SetValue("U_IssueWt", oClick_ColRow - 1, oMat01.Columns.Item("PQty").Cells.Item(oClick_ColRow).Specific.Value);
                                }
                            }
                            else
                            {
                                //원재료 단중으로 계산
                                oDS_PS_MM095L.SetValue("U_IssueQty", oClick_ColRow - 1, "0");
                                oDS_PS_MM095L.SetValue("U_IssueWt", oClick_ColRow - 1, Math.Round(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(oClick_ColRow).Specific.Value) * Convert.ToDouble(oMat02.Columns.Item("UnWeight").Cells.Item(pVal.Row).Specific.Value), 2));
                            }
                            oMat01.LoadFromDataSource();
                            oForm.Freeze(false);
                            //해당라인으로 이동
                            if (oClick_ColRow == oMat01.RowCount)
                            {
                                oMat01.SelectRow(oClick_ColRow, true, false);
                            }
                            else
                            {
                                oMat01.SelectRow(oClick_ColRow + 1, true, false);
                            }
                        }
                        else
                        {
                            errMessage = "원자재 재고 데이터가 불확실 합니다. 관리자에게 문의 바랍니다.";
                            throw new Exception();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
               if(errMessage != string.Empty)
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
                    oForm.Items.Item("Mat02").Top = oForm.Items.Item("Comments").Top - 105;
                    oForm.Items.Item("Mat02").Left = 6;
                    oForm.Items.Item("Mat02").Width = oForm.Width - 18;
                    oForm.Items.Item("Mat02").Height = 90;
                    oMat02.Columns.Item("ItemCode").Width = 100;
                    oMat02.Columns.Item("ItemName").Width = 200;
                    oMat02.Columns.Item("WhsCode").Width = 100;
                    oMat02.Columns.Item("WhsName").Width = 100;
                    oMat02.Columns.Item("Qty").Width = 100;
                    oMat02.Columns.Item("Weight").Width = 100;

                    oForm.Items.Item("Mat01").Top = 61;
                    oForm.Items.Item("Mat01").Left = 6;
                    oForm.Items.Item("Mat01").Width = oForm.Width - 18;
                    oForm.Items.Item("Mat01").Height = oForm.Height - 280;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            int i;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            if (oDS_PS_MM095H.GetValue("Canceled", 0).ToString().Trim() == "N")
                            {
                                if (PSH_Globals.SBO_Application.MessageBox("해당 문서가 취소됩니다 계속하시겠습니까??", 1, "&확인", "&취소") != 1)
                                {
                                    BubbleEvent = false;
                                    return;
                                }

                                if (PS_MM095_Add_InventoryGenEntry() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                dataHelpClass.MDC_GF_Message( "취소된 생산실적입니다. 확인하세요.",  "E");
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            PS_MM095_FormItemEnabled();
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }

                                oMat01.FlushToDataSource();
                                oDS_PS_MM095L.RemoveRecord(oDS_PS_MM095L.Size - 1);
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();
                            }
                            break;
                        case "1281": //찾기
                            PS_MM095_FormItemEnabled();
                            oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            oForm.Freeze(true);
                            PS_MM095_Initialization();
                            PS_MM095_FormItemEnabled();
                            PS_MM095_FormClear();
                            oForm.Items.Item("BPLId").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                            oForm.Items.Item("OrdGbn").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                            oDS_PS_MM095H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            oForm.Freeze(false);
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                        case "1287": //복제
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }
    }
}

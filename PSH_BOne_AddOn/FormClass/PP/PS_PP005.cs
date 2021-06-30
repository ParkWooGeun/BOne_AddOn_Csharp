using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 제품 원재료관계등록
    /// </summary>
    internal class PS_PP005 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP005H; //등록헤더
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP005.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP005_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP005");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);
                PS_PP005_CreateItems();
                PS_PP005_SetComboBox();
                
                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1287", false); // 복제
                oForm.EnableMenu("1286", false); // 닫기
                oForm.EnableMenu("1284", false); // 취소
                oForm.EnableMenu("1293", false); // 행삭제
                oForm.EnableMenu("1282", true);
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
        private void PS_PP005_CreateItems()
        {
            try
            {
                oDS_PS_PP005H = oForm.DataSources.DBDataSources.Item("@PS_PP005H");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("ItmBSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("ItmBSort").Specific.DataBind.SetBound(true, "", "ItmBSort");

                oForm.DataSources.UserDataSources.Add("ItmMSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("ItmMSort").Specific.DataBind.SetBound(true, "", "ItmMSort");

                oForm.DataSources.UserDataSources.Add("ItemCod1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("ItemCod1").Specific.DataBind.SetBound(true, "", "ItemCod1");

                oForm.DataSources.UserDataSources.Add("ItemNam1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("ItemNam1").Specific.DataBind.SetBound(true, "", "ItemNam1");

                oForm.DataSources.UserDataSources.Add("ItemCod2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("ItemCod2").Specific.DataBind.SetBound(true, "", "ItemCod2");

                oForm.DataSources.UserDataSources.Add("ItemNam2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("ItemNam2").Specific.DataBind.SetBound(true, "", "ItemNam2");

                oForm.DataSources.UserDataSources.Add("UnWeight", SAPbouiCOM.BoDataType.dt_PERCENT, 10);
                oForm.Items.Item("UnWeight").Specific.DataBind.SetBound(true, "", "UnWeight");

                oForm.DataSources.UserDataSources.Add("BaseChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("BaseChk").Specific.DataBind.SetBound(true, "", "BaseChk");
                oForm.Items.Item("BaseChk").Specific.Checked = false;

                oForm.DataSources.UserDataSources.Add("ConvChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ConvChk").Specific.DataBind.SetBound(true, "", "ConvChk");
                oForm.Items.Item("ConvChk").Specific.Checked = false;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP005_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBSort").Specific, "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code", "", false, false);
                oForm.Items.Item("ItmBSort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FlushToItemValue
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_PP005_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                switch (oUID)
                {
                    case "Mat01":
                        if (oCol == "ItemCod2")
                        {
                            sQry = "Select ItemName From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCod2").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ItemNam2").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();
                        }
                        break;
                    case "ItemCod1":
                        sQry = "Select ItemName, ItmMSort = U_ItmMsort From OITM Where ItemCode = '" + oForm.Items.Item("ItemCod1").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("ItemNam1").Specific.String = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();

                        if (oRecordSet01.RecordCount > 0)
                        {
                            oForm.Items.Item("ItmMSort").Specific.Select(oRecordSet01.Fields.Item("ItmMSort").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                        break;
                    case "ItemCod2":
                        sQry = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item("ItemCod2").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("ItemNam2").Specific.String = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();
                        break;
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
        /// 데이터 insert
        /// </summary>
        /// <returns></returns>
        private bool PS_PP005_AddData()
        {
            bool returnValue = false;
            string sQry;
            string DocEntry;
            double UnWeight;
            string ItemNam2;
            string ItemNam1;
            string ItemCod1;
            string ItemCod2;
            string Indate;
            string baseChk;
            string convChk;
            string errMessage = string.Empty;

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ItemCod1 = oForm.Items.Item("ItemCod1").Specific.Value.ToString().Trim();
                ItemNam1 = oForm.Items.Item("ItemNam1").Specific.Value.ToString().Trim();
                ItemCod2 = oForm.Items.Item("ItemCod2").Specific.Value.ToString().Trim();
                ItemNam2 = oForm.Items.Item("ItemNam2").Specific.Value.ToString().Trim();
                UnWeight = Convert.ToDouble(oForm.Items.Item("UnWeight").Specific.Value.ToString().Trim());
                baseChk = oForm.Items.Item("BaseChk").Specific.Checked == true ? "Y" : "N";
                convChk = oForm.Items.Item("ConvChk").Specific.Checked == true ? "Y" : "N";
                Indate = DateTime.Now.ToString("yyyyMMdd");

                sQry = "Select U_ItemCod1, U_ItemCod2 From [@PS_PP005H] Where U_ItemCod1 ='" + ItemCod1 + "' AND U_ItemCod2 = '" + ItemCod2 + "'";
                RecordSet01.DoQuery(sQry);

                if (RecordSet01.RecordCount > 0)
                {
                    errMessage = "기존자료가 존재합니다.";
                    throw new Exception();
                }

                if (UnWeight <= 0)
                {
                    errMessage = "단조품은 1, 중량단중은 개당 단중을 입력하여야 합니다.";
                    throw new Exception();
                }

                sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_PP005H]";
                RecordSet01.DoQuery(sQry);
                DocEntry = Convert.ToString(Convert.ToDouble(RecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1);

                sQry = "INSERT INTO [@PS_PP005H]";
                sQry += " (";
                sQry += " DocEntry,";
                sQry += " DocNum,";
                sQry += " U_ItemCod1,";
                sQry += " U_ItemNam1,";
                sQry += " U_ItemCod2,";
                sQry += " U_ItemNam2,";
                sQry += " U_UnWeight,";
                sQry += " U_InDate,";
                sQry += " U_BaseChk,";
                sQry += " U_ConvChk";
                sQry += " ) ";
                sQry += "VALUES(";
                sQry += DocEntry + ",";
                sQry += DocEntry + ",";
                sQry += "'" + ItemCod1 + "',";
                sQry += "'" + ItemNam1 + "',";
                sQry += "'" + ItemCod2 + "',";
                sQry += "'" + ItemNam2 + "',";
                sQry += UnWeight + ",";
                sQry += "'" + Indate + "',";
                sQry += "'" + baseChk + "',";
                sQry += "'" + convChk + "'";
                sQry += ")";
                RecordSet01.DoQuery(sQry);

                PSH_Globals.SBO_Application.StatusBar.SetText("제품코드 및 원자재코드 정상등록!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

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
        /// 데이터 Update
        /// </summary>
        /// <returns></returns>
        private bool PS_PP005_UpdateData()
        {
            bool returnValue = false;
            short i;
            string sQry;
            string DocEntry;
            string ItemCod2;
            string ItemNam2;
            string MoDate;
            double UnWeight;
            string baseChk;
            string convChk;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                MoDate = DateTime.Now.ToString("yyyyMMdd");

                oMat01.FlushToDataSource();

                for (i = 1; i <= oMat01.RowCount; i++)
                {
                    if (oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked == true)
                    {
                        UnWeight = Convert.ToDouble(oMat01.Columns.Item("UnWeight").Cells.Item(i).Specific.Value);

                        if (UnWeight <= 0)
                        {
                            errMessage = "단조품은 1, 중량단중은 개당 단중을 입력하여야 합니다.";
                            throw new Exception();
                        }
                    }
                }

                for (i = 1; i <= oMat01.RowCount; i++)
                {
                    if (oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked == true)
                    {
                        DocEntry = oMat01.Columns.Item("DocEntry").Cells.Item(i).Specific.Value;
                        ItemCod2 = oMat01.Columns.Item("ItemCod2").Cells.Item(i).Specific.Value;
                        ItemNam2 = oMat01.Columns.Item("ItemNam2").Cells.Item(i).Specific.Value;
                        UnWeight = Convert.ToDouble(oMat01.Columns.Item("UnWeight").Cells.Item(i).Specific.Value);
                        baseChk = oMat01.Columns.Item("BaseChk").Cells.Item(i).Specific.Checked == true ? "Y" : "N";
                        convChk = oMat01.Columns.Item("ConvChk").Cells.Item(i).Specific.Checked == true ? "Y" : "N";

                        sQry = "Update [@PS_PP005H] set ";
                        sQry += " U_ItemCod2   = '" + ItemCod2 + "',";
                        sQry += " U_ItemNam2   = '" + ItemNam2 + "',";
                        sQry += " U_UnWeight   = " + UnWeight + ",";
                        sQry += " U_BaseChk   = '" + baseChk + "',";
                        sQry += " U_ConvChk   = '" + convChk + "',";
                        sQry += " U_MoDate     = '" + MoDate + "'";
                        sQry += " Where DocEntry = '" + DocEntry + "'";
                        RecordSet01.DoQuery(sQry);
                    }
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("원자재코드 수정완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

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
        /// 데이터 Delete
        /// </summary>
        private void PS_PP005_DeleteData()
        {
            short i;
            string sQry;
            string DocEntry;
            string ItemCod2;
            string ItemNam2;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();

                for (i = 1; i <= oMat01.RowCount; i++)
                {
                    if (oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked == true)
                    {
                        DocEntry = oMat01.Columns.Item("DocEntry").Cells.Item(i).Specific.Value;
                        ItemCod2 = oMat01.Columns.Item("ItemCod2").Cells.Item(i).Specific.Value;
                        ItemNam2 = oMat01.Columns.Item("ItemNam2").Cells.Item(i).Specific.Value;

                        sQry = "Delete From [@PS_PP005H] where DocEntry = '" + DocEntry + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                PS_PP005_LoadData();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 조회데이터 가져오기
        /// </summary>
        private void PS_PP005_LoadData()
        {
            short i;
            string sQry;
            string ItmBsort;
            string ItemCod1;
            string ItemCod2;
            string ItmMsort;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                ItemCod1 = oForm.Items.Item("ItemCod1").Specific.Value.ToString().Trim();
                ItemCod2 = oForm.Items.Item("ItemCod2").Specific.Value.ToString().Trim();
                ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
                ItmMsort = oForm.Items.Item("ItmMSort").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(ItemCod1))
                {
                    ItemCod1 = "%";
                }
                    
                if (string.IsNullOrEmpty(ItemCod2))
                {
                    ItemCod2 = "%";
                }
                    
                if (string.IsNullOrEmpty(ItmMsort))
                {
                    ItmMsort = "%";
                }
                    
                if (string.IsNullOrEmpty(ItmBsort))
                {
                    ItmBsort = "%";
                }

                sQry = "EXEC [PS_PP005_01] '" + ItmBsort + "','" + ItmMsort + "','" + ItemCod1 + "','" + ItemCod2 + "'";

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_PP005H.Clear();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_PP005H.Size)
                    {
                        oDS_PS_PP005H.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_PP005H.Offset = i;

                    oDS_PS_PP005H.SetValue("DocEntry", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                    oDS_PS_PP005H.SetValue("DocNum", i, Convert.ToString(i + 1));
                    oDS_PS_PP005H.SetValue("U_ItemCod1", i, oRecordSet01.Fields.Item("U_ItemCod1").Value.ToString().Trim());
                    oDS_PS_PP005H.SetValue("U_ItemNam1", i, oRecordSet01.Fields.Item("U_ItemNam1").Value.ToString().Trim());
                    oDS_PS_PP005H.SetValue("U_ItmMSort", i, oRecordSet01.Fields.Item("U_ItmMSort").Value.ToString().Trim());
                    oDS_PS_PP005H.SetValue("U_BaseChk", i, oRecordSet01.Fields.Item("U_BaseChk").Value.ToString().Trim());
                    oDS_PS_PP005H.SetValue("U_ConvChk", i, oRecordSet01.Fields.Item("U_ConvChk").Value.ToString().Trim());
                    oDS_PS_PP005H.SetValue("U_ItemCod2", i, oRecordSet01.Fields.Item("U_ItemCod2").Value.ToString().Trim());
                    oDS_PS_PP005H.SetValue("U_ItemNam2", i, oRecordSet01.Fields.Item("U_ItemNam2").Value.ToString().Trim());
                    oDS_PS_PP005H.SetValue("U_UnWeight", i, oRecordSet01.Fields.Item("U_UnWeight").Value.ToString().Trim());
                    oDS_PS_PP005H.SetValue("U_InDate", i, oRecordSet01.Fields.Item("U_InDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_PP005H.SetValue("U_MoDate", i, oRecordSet01.Fields.Item("U_MoDate").Value.ToString("yyyyMMdd"));

                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
            }
        }

        /// <summary>
        /// Header 필수 입력 사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP005_CheckHeaderDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("ItemCod1").Specific.Value.ToString().Trim()))
                {
                    errMessage = "제품코드는 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ItemNam1").Specific.Value.ToString().Trim()))
                {
                    errMessage = "제품코드명은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ItemCod2").Specific.Value.ToString().Trim()))
                {
                    errMessage = "원자재코드는 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ItemNam2").Specific.Value.ToString().Trim()))
                {
                    errMessage = "원자재명은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
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
        /// Line 필수 입력 사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP005_CheckLineDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            
            try
            {
                oMat01.FlushToDataSource();

                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                oMat01.LoadFromDataSource();
                
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
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "Btn01") //저장
                    {
                        if (PS_PP005_CheckHeaderDataValid() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }

                        if (PS_PP005_AddData() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }

                        oMat01.Clear();
                        oMat01.FlushToDataSource();
                        oMat01.LoadFromDataSource();
                    }
                    else if (pVal.ItemUID == "Btn02") //조회
                    {
                        PS_PP005_LoadData();
                    }
                    else if (pVal.ItemUID == "Btn03") //삭제
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("정말로 삭제하시겠습니까?", 1, "Yes", "No") == 1)
                        {
                            PS_PP005_DeleteData();
                        }
                        else
                        {
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "Btn04") //수정
                    {
                        PS_PP005_UpdateData();
                    }
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
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "ItemCod1")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ItemCod1").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "ItemCod2")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ItemCod2").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ColUID == "ItemCod2")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCod2").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
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
            string sQry;
            int sCount;
            int sSeq;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "ItmBSort")
                    {
                        sCount = Convert.ToInt32(oForm.Items.Item("ItmMSort").Specific.ValidValues.Count);
                        sSeq = sCount;
                        for (int i = 1; i <= sCount; i++)
                        {
                            oForm.Items.Item("ItmMSort").Specific.ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                            sSeq -= 1;
                        }

                        sQry = "SELECT U_Code, U_CodeName From [@PSH_ITMMSORT] Where U_rCode = '" + oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim() + "' Order by Code";
                        oRecordSet01.DoQuery(sQry);
                        while (!oRecordSet01.EoF)
                        {
                            oForm.Items.Item("ItmMSort").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                            oRecordSet01.MoveNext();
                        }
                        oForm.Items.Item("ItmMSort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
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
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "ItemCod1")
                        {
                            PS_PP005_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "ItemCod2")
                        {
                            PS_PP005_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ItemCod2")
                        {
                            PS_PP005_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                        }
                    }
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP005H);
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
        /// FORM_RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oMat01.AutoResizeColumns();
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
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            oMat01.Clear();
                            oDS_PS_PP005H.Clear();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
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

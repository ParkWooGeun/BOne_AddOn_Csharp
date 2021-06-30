//using System;
//using SAPbouiCOM;
//using PSH_BOne_AddOn.Data;
//using PSH_BOne_AddOn.Code;

//namespace PSH_BOne_AddOn
//{
//    /// <summary>
//    /// 제품 원재료관계등록
//    /// </summary>
//    internal class PS_PP005 : PSH_BaseClass
//    {
//        private string oFormUniqueID;
//        private SAPbouiCOM.Matrix oMat01;
//        private SAPbouiCOM.DBDataSource oDS_PS_PP005H; //등록헤더
//        private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
//        private string oLast_Col_UID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//        private int oLast_Col_Row; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//        private int oLast_Mode;

//        /// <summary>
//		/// Form 호출
//		/// </summary>
//		/// <param name="oFormDocEntry"></param>
//        public override void LoadForm(string oFormDocEntry)
//        {
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            try
//            {
//                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP005.srf");
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

//                //매트릭스의 타이틀높이와 셀높이를 고정
//                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//                {
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//                }

//                oFormUniqueID = "PS_PP005_" + SubMain.Get_TotalFormsCount();
//                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP005");

//                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
//                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

//                oForm.SupportedModes = -1;
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                //oForm.DataBrowser.BrowseBy = "DocNum";

//                oForm.Freeze(true);
//                PS_PP005_CreateItems();
//                PS_PP005_SetComboBox();

//                oForm.EnableMenu("1283", false); // 삭제
//                oForm.EnableMenu("1287", false); // 복제
//                oForm.EnableMenu("1286", true); // 닫기
//                oForm.EnableMenu("1284", true); // 취소
//                oForm.EnableMenu("1293", true); // 행삭제
//                oForm.EnableMenu("1282", true);
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Update();
//                oForm.Freeze(false);
//                oForm.Visible = true;
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
//            }
//        }

//        /// <summary>
//        /// 화면 Item 생성
//        /// </summary>
//        private void PS_PP005_CreateItems()
//        {
//            try
//            {
//                oDS_PS_PP005H = oForm.DataSources.DBDataSources.Item("@PS_PP005H");
//                oMat01 = oForm.Items.Item("Mat01").Specific;
//                oMat01.AutoResizeColumns();

//                oForm.DataSources.UserDataSources.Add("ItmBSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//                oForm.Items.Item("ItmBSort").Specific.DataBind.SetBound(true, "", "ItmBSort");

//                oForm.DataSources.UserDataSources.Add("ItmMSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//                oForm.Items.Item("ItmMSort").Specific.DataBind.SetBound(true, "", "ItmMSort");

//                oForm.DataSources.UserDataSources.Add("ItemCod1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//                oForm.Items.Item("ItemCod1").Specific.DataBind.SetBound(true, "", "ItemCod1");

//                oForm.DataSources.UserDataSources.Add("ItemNam1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
//                oForm.Items.Item("ItemNam1").Specific.DataBind.SetBound(true, "", "ItemNam1");

//                oForm.DataSources.UserDataSources.Add("ItemCod2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//                oForm.Items.Item("ItemCod2").Specific.DataBind.SetBound(true, "", "ItemCod2");

//                oForm.DataSources.UserDataSources.Add("ItemNam2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
//                oForm.Items.Item("ItemNam2").Specific.DataBind.SetBound(true, "", "ItemNam2");

//                oForm.DataSources.UserDataSources.Add("UnWeight", SAPbouiCOM.BoDataType.dt_PERCENT, 10);
//                oForm.Items.Item("UnWeight").Specific.DataBind.SetBound(true, "", "UnWeight");

//                oForm.DataSources.UserDataSources.Add("BaseChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oForm.Items.Item("BaseChk").Specific.DataBind.SetBound(true, "", "BaseChk");
//                oForm.Items.Item("BaseChk").Specific.Checked = false;

//                oForm.DataSources.UserDataSources.Add("ConvChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oForm.Items.Item("ConvChk").Specific.DataBind.SetBound(true, "", "ConvChk");
//                oForm.Items.Item("ConvChk").Specific.Checked = false;
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// Combobox 설정
//        /// </summary>
//        private void PS_PP005_SetComboBox()
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBSort").Specific, "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code", "", false, false);
//                oForm.Items.Item("ItmBSort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// FlushToItemValue
//        /// </summary>
//        /// <param name="oUID"></param>
//        /// <param name="oRow"></param>
//        /// <param name="oCol"></param>
//        private void PS_PP005_FlushToItemValue(string oUID, int oRow, string oCol)
//        {
//            string sQry;
//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                switch (oUID)
//                {
//                    case "Mat01":
//                        if (oCol == "ItemCod2")
//                        {
//                            sQry = "Select ItemName  From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCod2").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
//                            oRecordSet01.DoQuery(sQry);
//                            oMat01.Columns.Item("ItemNam2").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();
//                        }
//                        break;
//                    case "ItemCod1":
//                        sQry = "Select ItemName, ItmMSort = U_ItmMsort  From OITM Where ItemCode = '" + oForm.Items.Item("ItemCod1").Specific.Value.ToString().Trim() + "'";
//                        oRecordSet01.DoQuery(sQry);
//                        oForm.Items.Item("ItemNam1").Specific.String = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();
//                        oForm.Items.Item("ItmMSort").Specific.Select(oRecordSet01.Fields.Item("ItmMSort").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
//                        break;
//                    case "ItemCod2":
//                        sQry = "Select ItemName  From OITM Where ItemCode = '" + oForm.Items.Item("ItemCod2").Specific.Value.ToString().Trim() + "'";
//                        oRecordSet01.DoQuery(sQry);
//                        oForm.Items.Item("ItemNam2").Specific.String = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();
//                        break;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
//            }
//        }

//        /// <summary>
//        /// 데이터 insert
//        /// </summary>
//        /// <returns></returns>
//        private bool PS_PP005_AddData()
//        {
//            bool returnValue = false;
//            string sQry;
//            string DocEntry;
//            double UnWeight;
//            string ItemNam2;
//            string ItemNam1;
//            string ItemCod1;
//            string ItemCod2;
//            string Indate;
//            string baseChk;
//            string convChk;
//            string errMessage = string.Empty;

//            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                ItemCod1 = oForm.Items.Item("ItemCod1").Specific.Value.ToString().Trim();
//                ItemNam1 = oForm.Items.Item("ItemNam1").Specific.Value.ToString().Trim();
//                ItemCod2 = oForm.Items.Item("ItemCod2").Specific.Value.ToString().Trim();
//                ItemNam2 = oForm.Items.Item("ItemNam2").Specific.Value.ToString().Trim();
//                UnWeight = Convert.ToDouble(oForm.Items.Item("UnWeight").Specific.Value.ToString().Trim());
//                baseChk = oForm.Items.Item("BaseChk").Specific.Checked == true ? "Y" : "N";
//                convChk = oForm.Items.Item("ConvChk").Specific.Checked == true ? "Y" : "N";
//                Indate = DateTime.Now.ToString("yyyyMMdd");

//                sQry = "Select U_ItemCod1, U_ItemCod2 From [@PS_PP005H] Where U_ItemCod1 ='" + ItemCod1 + "' AND U_ItemCod2 = '" + ItemCod2 + "'";
//                RecordSet01.DoQuery(sQry);

//                if (RecordSet01.RecordCount > 0)
//                {
//                    errMessage = "기존자료가 존재합니다.";
//                    throw new Exception();
//                }

//                if (UnWeight <= 0)
//                {
//                    errMessage = "단조품은 1, 중량단중은 개당 단중을 입력하여야 합니다.";
//                    throw new Exception();
//                }

//                sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_PP005H]";
//                RecordSet01.DoQuery(sQry);
//                DocEntry = Convert.ToString(Convert.ToDouble(RecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1);

//                sQry = "INSERT INTO [@PS_PP005H]";
//                sQry += " (";
//                sQry += " DocEntry,";
//                sQry += " DocNum,";
//                sQry += " U_ItemCod1,";
//                sQry += " U_ItemNam1,";
//                sQry += " U_ItemCod2,";
//                sQry += " U_ItemNam2,";
//                sQry += " U_UnWeight,";
//                sQry += " U_InDate,";
//                sQry += " U_BaseChk,";
//                sQry += " U_ConvChk";
//                sQry += " ) ";
//                sQry += "VALUES(";
//                sQry += DocEntry + ",";
//                sQry += DocEntry + ",";
//                sQry += "'" + ItemCod1 + "',";
//                sQry += "'" + ItemNam1 + "',";
//                sQry += "'" + ItemCod2 + "',";
//                sQry += "'" + ItemNam2 + "',";
//                sQry += UnWeight + ",";
//                sQry += Indate + ",";
//                sQry += "'" + baseChk + "',";
//                sQry += "'" + convChk + "'";
//                sQry += ")";
//                RecordSet01.DoQuery(sQry);

//                PSH_Globals.SBO_Application.StatusBar.SetText("제품코드 및 원자재코드 정상등록!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

//                returnValue = true;
//            }
//            catch (Exception ex)
//            {
//                if (errMessage != string.Empty)
//                {
//                    PSH_Globals.SBO_Application.MessageBox(errMessage);
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
//                }
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
//            }

//            return returnValue;
//        }

//        /// <summary>
//        /// 데이터 Update
//        /// </summary>
//        /// <returns></returns>
//        private bool PS_PP005_UpdateData()
//        {
//            bool returnValue = false;
//            short i;
//            string sQry;
//            string DocEntry;
//            string ItemCod2;
//            string ItemNam2;
//            string Chk;
//            string MoDate;
//            double UnWeight;
//            string baseChk;
//            string convChk;
//            string errMessage = string.Empty;
//            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                MoDate = DateTime.Now.ToString("yyyyMMdd");

//                oMat01.FlushToDataSource();

//                for (i = 1; i <= oMat01.RowCount; i++)
//                {
//                    Chk = oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked;

//                    if (Convert.ToBoolean(Chk) == true)
//                    {
//                        UnWeight = oMat01.Columns.Item("UnWeight").Cells.Item(i).Specific.Value;

//                        if (UnWeight <= 0)
//                        {
//                            errMessage = "단조품은 1, 중량단중은 개당 단중을 입력하여야 합니다.";
//                            throw new Exception();
//                        }
//                    }
//                }

//                for (i = 1; i <= oMat01.RowCount; i++)
//                {
//                    Chk = oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked;

//                    if (Convert.ToBoolean(Chk) == true)
//                    {
//                        DocEntry = oMat01.Columns.Item("DocEntry").Cells.Item(i).Specific.Value;
//                        ItemCod2 = oMat01.Columns.Item("ItemCod2").Cells.Item(i).Specific.Value;
//                        ItemNam2 = oMat01.Columns.Item("ItemNam2").Cells.Item(i).Specific.Value;
//                        UnWeight = Convert.ToDouble(oMat01.Columns.Item("UnWeight").Cells.Item(i).Specific.Value);
//                        baseChk = oMat01.Columns.Item("BaseChk").Cells.Item(i).Specific.Checked == true ? "Y" : "N";
//                        convChk = oMat01.Columns.Item("ConvChk").Cells.Item(i).Specific.Checked == true ? "Y" : "N";

//                        sQry = "Update [@PS_PP005H] set ";
//                        sQry += " U_ItemCod2   = '" + ItemCod2 + "',";
//                        sQry += " U_ItemNam2   = '" + ItemNam2 + "',";
//                        sQry += " U_UnWeight   = " + UnWeight + ",";
//                        sQry += " U_BaseChk   = '" + baseChk + "',";
//                        sQry += " U_ConvChk   = '" + convChk + "',";
//                        sQry += " U_MoDate     = '" + MoDate + "'";
//                        sQry += " Where DocEntry = '" + DocEntry + "'";
//                        RecordSet01.DoQuery(sQry);
//                    }
//                }

//                PSH_Globals.SBO_Application.StatusBar.SetText("원자재코드 수정완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

//                returnValue = true;
//            }
//            catch (Exception ex)
//            {
//                if (errMessage != string.Empty)
//                {
//                    PSH_Globals.SBO_Application.MessageBox(errMessage);
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
//                }
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
//            }

//            return returnValue;
//        }

//        /// <summary>
//        /// 데이터 Delete
//        /// </summary>
//        private void PS_PP005_DeleteData()
//        {
//            short i;
//            string sQry;
//            string DocEntry;
//            string ItemCod2;
//            string ItemNam2;
//            string Chk;
//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                oMat01.FlushToDataSource();

//                for (i = 1; i <= oMat01.RowCount; i++)
//                {
//                    Chk = oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked;

//                    if (Convert.ToBoolean(Chk) == true)
//                    {
//                        DocEntry = oMat01.Columns.Item("DocEntry").Cells.Item(i).Specific.Value;
//                        ItemCod2 = oMat01.Columns.Item("ItemCod2").Cells.Item(i).Specific.Value;
//                        ItemNam2 = oMat01.Columns.Item("ItemNam2").Cells.Item(i).Specific.Value;

//                        sQry = "Delete From [@PS_PP005H] where DocEntry = '" + DocEntry + "'";
//                        oRecordSet01.DoQuery(sQry);
//                    }
//                }

//                oMat01.Clear();
//                oMat01.FlushToDataSource();
//                oMat01.LoadFromDataSource();
//                PS_PP005_LoadData();
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
//            }

//        }

//        /// <summary>
//        /// 조회데이터 가져오기
//        /// </summary>
//        private void PS_PP005_LoadData()
//        {
//            short i;
//            string sQry;
//            string ItmBsort;
//            string ItemCod1;
//            string ItemCod2;
//            string ItmMsort;
//            string errMessage = string.Empty;
//            SAPbouiCOM.ProgressBar ProgBar01 = null;
//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                oForm.Freeze(true);

//                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

//                ItemCod1 = oForm.Items.Item("ItemCod1").Specific.Value.ToString().Trim();
//                ItemCod2 = oForm.Items.Item("ItemCod2").Specific.Value.ToString().Trim();
//                ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
//                ItmMsort = oForm.Items.Item("ItmMSort").Specific.Value.ToString().Trim();

//                if (string.IsNullOrEmpty(ItemCod1))
//                {
//                    ItemCod1 = "%";
//                }

//                if (string.IsNullOrEmpty(ItemCod2))
//                {
//                    ItemCod2 = "%";
//                }

//                if (string.IsNullOrEmpty(ItmMsort))
//                {
//                    ItmMsort = "%";
//                }

//                if (string.IsNullOrEmpty(ItmBsort))
//                {
//                    ItmBsort = "%";
//                }

//                sQry = "EXEC [PS_PP005_01] '" + ItmBsort + "','" + ItmMsort + "','" + ItemCod1 + "','" + ItemCod2 + "'";

//                oRecordSet01.DoQuery(sQry);

//                oMat01.Clear();
//                oDS_PS_PP005H.Clear();
//                oMat01.LoadFromDataSource();

//                if (oRecordSet01.RecordCount == 0)
//                {
//                    errMessage = "조회 결과가 없습니다. 확인하세요.";
//                    throw new Exception();
//                }

//                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
//                {
//                    if (i + 1 > oDS_PS_PP005H.Size)
//                    {
//                        oDS_PS_PP005H.InsertRecord(i);
//                    }

//                    oMat01.AddRow();
//                    oDS_PS_PP005H.Offset = i;

//                    oDS_PS_PP005H.SetValue("DocEntry", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
//                    oDS_PS_PP005H.SetValue("DocNum", i, Convert.ToString(i + 1));
//                    oDS_PS_PP005H.SetValue("U_ItemCod1", i, oRecordSet01.Fields.Item("U_ItemCod1").Value.ToString().Trim());
//                    oDS_PS_PP005H.SetValue("U_ItemNam1", i, oRecordSet01.Fields.Item("U_ItemNam1").Value.ToString().Trim());
//                    oDS_PS_PP005H.SetValue("U_ItmMSort", i, oRecordSet01.Fields.Item("U_ItmMSort").Value.ToString().Trim());
//                    oDS_PS_PP005H.SetValue("U_BaseChk", i, oRecordSet01.Fields.Item("U_BaseChk").Value.ToString().Trim());
//                    oDS_PS_PP005H.SetValue("U_ConvChk", i, oRecordSet01.Fields.Item("U_ConvChk").Value.ToString().Trim());
//                    oDS_PS_PP005H.SetValue("U_ItemCod2", i, oRecordSet01.Fields.Item("U_ItemCod2").Value.ToString().Trim());
//                    oDS_PS_PP005H.SetValue("U_ItemNam2", i, oRecordSet01.Fields.Item("U_ItemNam2").Value.ToString().Trim());
//                    oDS_PS_PP005H.SetValue("U_UnWeight", i, oRecordSet01.Fields.Item("U_UnWeight").Value.ToString().Trim());
//                    oDS_PS_PP005H.SetValue("U_InDate", i, oRecordSet01.Fields.Item("U_InDate").Value.ToString("yyyyMMdd"));
//                    oDS_PS_PP005H.SetValue("U_MoDate", i, oRecordSet01.Fields.Item("U_MoDate").Value.ToString("yyyyMMdd"));

//                    oRecordSet01.MoveNext();
//                    ProgBar01.Value += 1;
//                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
//                }
//                oMat01.LoadFromDataSource();
//                oMat01.AutoResizeColumns();
//            }
//            catch (Exception ex)
//            {
//                if (errMessage != string.Empty)
//                {
//                    PSH_Globals.SBO_Application.MessageBox(errMessage);
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
//                }
//            }
//            finally
//            {
//                oForm.Freeze(false);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
//                if (ProgBar01 != null)
//                {
//                    ProgBar01.Stop();
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
//                }
//            }
//        }

//        #region PS_PP005_DelHeaderSpaceLine
//    }
//}

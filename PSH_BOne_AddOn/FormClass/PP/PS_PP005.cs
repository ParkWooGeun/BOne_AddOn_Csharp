using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

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
        private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLast_Col_UID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLast_Col_Row; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oLast_Mode;

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
                oForm.EnableMenu("1286", true); // 닫기
                oForm.EnableMenu("1284", true); // 취소
                oForm.EnableMenu("1293", true); // 행삭제
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

                oForm.DataSources.UserDataSources.Add("ItmBSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("ItmBSort").Specific.DataBind.SetBound(true, "", "ItmBSort");

                oForm.DataSources.UserDataSources.Add("ItmMSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("ItmMSort").Specific.DataBind.SetBound(true, "", "ItmMSort");

                oForm.DataSources.UserDataSources.Add("ItemCod1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ItemCod1").Specific.DataBind.SetBound(true, "", "ItemCod1");

                oForm.DataSources.UserDataSources.Add("ItemNam1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("ItemNam1").Specific.DataBind.SetBound(true, "", "ItemNam1");

                oForm.DataSources.UserDataSources.Add("ItemCod2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
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
                            sQry = "Select ItemName  From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCod2").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ItemNam2").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();
                        }
                        break;
                    case "ItemCod1":
                        sQry = "Select ItemName, ItmMSort = U_ItmMsort  From OITM Where ItemCode = '" + oForm.Items.Item("ItemCod1").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("ItemNam1").Specific.String = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();
                        oForm.Items.Item("ItmMSort").Specific.Select(oRecordSet01.Fields.Item("ItmMSort").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                        break;
                    case "ItemCod2":
                        sQry = "Select ItemName  From OITM Where ItemCode = '" + oForm.Items.Item("ItemCod2").Specific.Value.ToString().Trim() + "'";
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
                sQry += Indate + ",";
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
            string Chk;
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
                    Chk = oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked;

                    if (Convert.ToBoolean(Chk) == true)
                    {
                        UnWeight = oMat01.Columns.Item("UnWeight").Cells.Item(i).Specific.Value;

                        if (UnWeight <= 0)
                        {
                            errMessage = "단조품은 1, 중량단중은 개당 단중을 입력하여야 합니다.";
                            throw new Exception();
                        }
                    }
                }

                for (i = 1; i <= oMat01.RowCount; i++)
                {
                    Chk = oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked;

                    if (Convert.ToBoolean(Chk) == true)
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
            string Chk;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();

                for (i = 1; i <= oMat01.RowCount; i++)
                {
                    Chk = oMat01.Columns.Item("Chk").Cells.Item(i).Specific.Checked;

                    if (Convert.ToBoolean(Chk) == true)
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

        #region PS_PP005_DelHeaderSpaceLine
        //private bool PS_PP005_DelHeaderSpaceLine()
        //{
        //    bool returnValue = false;
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    short ErrNum = 0;
        //    string DocNum = null;

        //    ErrNum = 0;

        //    // 저장버튼 클릭시 필수입력 필드에 값이 있는지를 Check 한다.

        //    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    switch (true)
        //    {
        //        case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ItemCod1").Specific.Value)):
        //            ErrNum = 1;
        //            goto PS_PP005_DelHeaderSpaceLine_Error;
        //            break;
        //        case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ItemNam1").Specific.Value)):
        //            ErrNum = 2;
        //            goto PS_PP005_DelHeaderSpaceLine_Error;
        //            break;
        //        case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ItemCod2").Specific.Value)):
        //            ErrNum = 3;
        //            goto PS_PP005_DelHeaderSpaceLine_Error;
        //            break;
        //        case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ItemNam2").Specific.Value)):
        //            ErrNum = 4;
        //            goto PS_PP005_DelHeaderSpaceLine_Error;
        //            break;
        //    }

        //    returnValue = true;
        //    return returnValue;
        //PS_PP005_DelHeaderSpaceLine_Error:

        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //    if (ErrNum == 1)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "제품코드는 필수입력사항입니다. 확인하세요.", ref "E");
        //    }
        //    else if (ErrNum == 2)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "제품코드명은 필수입력사항입니다. 확인하세요.", ref "E");
        //    }
        //    else if (ErrNum == 3)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "원자재코드는 필수입력사항입니다. 확인하세요.", ref "E");
        //    }
        //    else if (ErrNum == 4)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "원자재명은 필수입력사항입니다. 확인하세요.", ref "E");
        //    }
        //    else
        //    {
        //        MDC_Com.MDC_GF_Message(ref "PS_PP005_DelHeaderSpaceLine_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //    }
        //    returnValue = false;
        //    return returnValue;
        //}
        #endregion

        #region PS_PP005_DelMatrixSpaceLine
        //private bool PS_PP005_DelMatrixSpaceLine()
        //{
        //    bool returnValue = false;
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;
        //    short ErrNum = 0;
        //    SAPbobsCOM.Recordset oRecordSet = null;
        //    string sQry = null;

        //    oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    ErrNum = 0;

        //    oMat01.FlushToDataSource();

        //    //// 라인
        //    if (oMat01.VisualRowCount == 0)
        //    {
        //        ErrNum = 1;
        //        goto PS_PP005_DelMatrixSpaceLine_Error;
        //    }

        //    oMat01.LoadFromDataSource();

        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    returnValue = true;
        //    return returnValue;
        //PS_PP005_DelMatrixSpaceLine_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    if (ErrNum == 1)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하세요.", ref "E");
        //    }
        //    else
        //    {
        //        MDC_Com.MDC_GF_Message(ref "PS_PP005_DelMatrixSpaceLine_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //    }
        //    returnValue = false;
        //    return returnValue;
        //}
        #endregion

        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{

        //    // ERROR: Not supported in C#: OnErrorStatement


        //    short i = 0;
        //    string sQry = null;
        //    SAPbobsCOM.Recordset oRecordSet01 = null;

        //    oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    int sCount = 0;
        //    int sSeq = 0;
        //    ////BeforeAction = True
        //    if ((pVal.BeforeAction == true))
        //    {
        //        switch (pVal.EventType)
        //        {
        //            //et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //                ////1                               '버튼 클릭시 발생하는 Event
        //                if (pVal.ItemUID == "Btn01")
        //                {
        //                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //                    {
        //                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //                        {
        //                            //저장 버튼클릭
        //                            if (pVal.ItemUID == "Btn01")
        //                            {
        //                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //                                {
        //                                    if (PS_PP005_DelHeaderSpaceLine() == false)
        //                                    {
        //                                        BubbleEvent = false;
        //                                        return;
        //                                    }

        //                                    if (PS_PP005_AddData() == false)
        //                                    {
        //                                        BubbleEvent = false;
        //                                        return;
        //                                    }

        //                                    oMat01.Clear();
        //                                    oMat01.FlushToDataSource();
        //                                    oMat01.LoadFromDataSource();
        //                                    PS_PP005_AddMatrixRow(0, ref true);

        //                                    //                        Call Delete_EmptyRow
        //                                    oLast_Mode = oForm.Mode;
        //                                    //                                ElseIf oForm.Mode = fm_UPDATE_MODE Then
        //                                    //                                    If Updatedata(pVal) = False Then
        //                                    //                                        BubbleEvent = False
        //                                    //                                        Exit Sub
        //                                    //                                    End If
        //                                    //                                    Call PS_PP005_LoadData
        //                                }
        //                            }
        //                        }
        //                        oLast_Mode = oForm.Mode;
        //                    }
        //                    //내역조회
        //                }
        //                else if (pVal.ItemUID == "Btn02")
        //                {
        //                    PS_PP005_LoadData();
        //                }
        //                else if (pVal.ItemUID == "Btn03")
        //                {
        //                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //                    ///fm_VIEW_MODE
        //                    PS_PP005_DeleteData();
        //                }
        //                else if (pVal.ItemUID == "Btn04")
        //                {
        //                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        //                    PS_PP005_UpdateData();
        //                }
        //                return;

        //                break;
        //            //et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //                ////2                         '탭키를 눌렀을 때 발생하는 Event

        //                // 질의 관리자 창을 사용할 때... 선언부분은 FlushTo ItemValue에서 선언

        //                if (pVal.CharPressed == 9)
        //                {
        //                    if (pVal.ItemUID == "ItemCod1")
        //                    {
        //                        //UPGRADE_WARNING: oForm.Items(ItemCod1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (string.IsNullOrEmpty(oForm.Items.Item("ItemCod1").Specific.Value))
        //                        {
        //                            SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //                            //질의 관리자 사용시(탭키를 눌렀을 때)
        //                            BubbleEvent = false;
        //                        }
        //                    }
        //                    else if (pVal.ItemUID == "ItemCod2")
        //                    {
        //                        //UPGRADE_WARNING: oForm.Items(ItemCod2).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (string.IsNullOrEmpty(oForm.Items.Item("ItemCod2").Specific.Value))
        //                        {
        //                            SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //                            BubbleEvent = false;
        //                        }
        //                        // Matrix에 질의 관리자 사용시 선언
        //                    }
        //                    else if (pVal.ColUID == "ItemCod2")
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns(ItemCod2).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCod2").Cells.Item(pVal.Row).Specific.Value))
        //                        {
        //                            SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //                            BubbleEvent = false;
        //                        }
        //                    }
        //                }
        //                break;

        //            //et_COMBO_SELECT ////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //                ////5
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CLICK:
        //                ////6
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //                ////7
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //                ////8
        //                break;
        //            //et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //                ////10
        //                break;


        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //                ////11
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //                ////18
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //                ////19
        //                break;
        //            //et_FORM_RESIZE//////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //                ////20
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //                ////27
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //                ////3
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //                ////4
        //                break;
        //            //et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //                ////17
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((pVal.BeforeAction == false))
        //    {
        //        switch (pVal.EventType)
        //        {
        //            //et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //                ////1
        //                if (pVal.ItemUID == "1")
        //                {
        //                    PS_PP005_AddMatrixRow(oMat01.RowCount, ref false);
        //                    oLast_Mode = 0;
        //                    //                    If oForm.Mode = fm_OK_MODE And oLast_Mode = fm_UPDATE_MODE Then
        //                    //
        //                    //                    End If
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //                ////2
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //                ////5

        //                if (pVal.ItemUID == "ItmBSort")
        //                {
        //                    //UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    sCount = Convert.ToInt32(Strings.Trim(oForm.Items.Item("ItmMSort").Specific.ValidValues.Count));
        //                    sSeq = sCount;
        //                    for (i = 1; i <= sCount; i++)
        //                    {
        //                        //UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oForm.Items.Item("ItmMSort").Specific.ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
        //                        sSeq = sSeq - 1;
        //                    }

        //                    ////중분류
        //                    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    sQry = "SELECT U_Code, U_CodeName From [@PSH_ITMMSORT] Where U_rCode = '" + Strings.Trim(oForm.Items.Item("ItmBSort").Specific.Value) + "' Order by Code";
        //                    oRecordSet01.DoQuery(sQry);
        //                    //                    oForm.Items("ItmMSort").Specific.ValidValues.Add "", ""
        //                    while (!(oRecordSet01.EoF))
        //                    {
        //                        //UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oForm.Items.Item("ItmMSort").Specific.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
        //                        oRecordSet01.MoveNext();
        //                    }
        //                    //                    If oRecordSet01.RecordCount <> 0 Then
        //                    //UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oForm.Items.Item("ItmMSort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                    //                    Else
        //                    //                        oForm.Items("ItmMSort").Specific.Select "", psk_ByValue
        //                    //                    End If
        //                }
        //                else if (pVal.ItemUID == "Mat01")
        //                {
        //                    ///                    If oForm.Mode = fm_ADD_MODE Then
        //                    ///                    Else
        //                    ///                        oForm.Mode = fm_UPDATE_MODE
        //                    ///                        Call LoadCaption
        //                    ///                    End If
        //                }
        //                break;

        //            case SAPbouiCOM.BoEventTypes.et_CLICK:
        //                ////6
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //                ////7
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //                ////8
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //                ////10       ' 필드의 값이 바뀌었을 때 동작하는 부분
        //                if (pVal.ItemChanged == true)
        //                {
        //                    if (pVal.ItemUID == "ItemCod1")
        //                    {
        //                        PS_PP005_FlushToItemValue(pVal.ItemUID);
        //                    }
        //                    if (pVal.ItemUID == "ItemCod2")
        //                    {
        //                        PS_PP005_FlushToItemValue(pVal.ItemUID);
        //                    }
        //                    if (pVal.ColUID == "ItemCod2")
        //                    {
        //                        PS_PP005_FlushToItemValue(pVal.ItemUID, ref pVal.Row, ref pVal.ColUID);
        //                    }
        //                }
        //                break;

        //            //                If pVal.ItemChanged = True Then
        //            //                    If pVal.ItemUID = "CardCode" Then
        //            //                        PS_PP005_FlushToItemValue pVal.ItemUID
        //            //                    ElseIf pVal.ItemUID = "CntcCode" Then
        //            //                        PS_PP005_FlushToItemValue pVal.ItemUID
        //            //                    ElseIf pVal.ItemUID = "Mat01" Then
        //            //                        If pVal.ColUID = "GADocLin" Then
        //            //                            PS_PP005_FlushToItemValue pVal.ItemUID, pVal.Row, pVal.ColUID
        //            //                        End If
        //            //                    End If
        //            //                End If

        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //                ////11
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //                ////18
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //                ////19
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //                ////20
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //                ////27
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //                ////3
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //                ////4
        //                break;
        //            //et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //                ////17
        //                SubMain.RemoveForms(oFormUniqueID);
        //                //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oForm = null;
        //                //UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oMat01 = null;
        //                break;
        //        }
        //    }
        //    return;
        //Raise_ItemEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;

        //    ////BeforeAction = True
        //    if ((pVal.BeforeAction == true))
        //    {
        //        switch (pVal.MenuUID)
        //        {
        //            case "1284":
        //                //취소
        //                break;


        //            case "1286":
        //                //닫기
        //                break;


        //            case "1293":
        //                //행삭제
        //                break;
        //            case "1281":
        //                //찾기
        //                break;
        //            case "1282":
        //                //추가
        //                break;
        //            case "1288":
        //            case "1289":
        //            case "1290":
        //            case "1291":
        //                //레코드이동버튼
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((pVal.BeforeAction == false))
        //    {
        //        switch (pVal.MenuUID)
        //        {
        //            //            Case "1284": '취소
        //            //                PS_PP005_EnableFormItem
        //            //                oForm.Items("DocNum").Click ct_Regular
        //            case "1286":
        //                //닫기
        //                break;
        //            case "1293":
        //                //행삭제

        //                if (oMat01.RowCount != oMat01.VisualRowCount)
        //                {
        //                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
        //                    }

        //                    oMat01.FlushToDataSource();
        //                    oDS_PS_PP005H.RemoveRecord(oDS_PS_PP005H.Size - 1);
        //                    //// Mat01에 마지막라인(빈라인) 삭제
        //                    oMat01.Clear();
        //                    oMat01.LoadFromDataSource();

        //                    //UPGRADE_WARNING: oMat01.Columns(PQDocNum).Cells(oMat01.RowCount).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("PQDocNum").Cells.Item(oMat01.RowCount).Specific.Value))
        //                    {
        //                        PS_PP005_AddMatrixRow(oMat01.RowCount, ref false);
        //                    }
        //                }
        //                break;

        //            case "1281":
        //                //찾기
        //                PS_PP005_EnableFormItem();
        //                break;
        //            //oForm.Items("DocNum").Click ct_Regular

        //            case "1282":
        //                //추가
        //                //                oForm.Items("ItemCod1").Specific.Value = ""
        //                //                oForm.Items("ItemCod2").Specific.Value = ""
        //                oMat01.Clear();
        //                oDS_PS_PP005H.Clear();
        //                break;
        //            //                oDS_PS_PP005H.GetValue("U_ItemCod1",0)

        //            case "1288":
        //            case "1289":
        //            case "1290":
        //            case "1291":
        //                //레코드이동버튼
        //                PS_PP005_EnableFormItem();
        //                if (oMat01.VisualRowCount > 0)
        //                {
        //                    //UPGRADE_WARNING: oMat01.Columns(PQDocNum).Cells(oMat01.VisualRowCount).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("PQDocNum").Cells.Item(oMat01.VisualRowCount).Specific.Value))
        //                    {
        //                        if (oDS_PS_PP005H.GetValue("Status", 0) == "O")
        //                        {
        //                            PS_PP005_AddMatrixRow(oMat01.RowCount, ref false);
        //                        }
        //                    }
        //                }
        //                break;

        //        }
        //    }
        //    return;
        //Raise_MenuEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    ////BeforeAction = True
        //    if ((BusinessObjectInfo.BeforeAction == true))
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                ////33
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                ////34
        //                break;
        //            //                If Add_oPurchaseOrders(2) = False Then
        //            //                    BubbleEvent = False
        //            //                    Exit Sub
        //            //                Else
        //            //                    Call Delete_EmptyRow
        //            //                End If

        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                ////35
        //                if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //                {
        //                    //                    If Update_oPurchaseOrders(2) = False Then
        //                    //                        oLast_Mode = 0
        //                    //                        BubbleEvent = False
        //                    //                        Exit Sub
        //                    //                    Else
        //                    //                        oLast_Mode = 0
        //                    //                        Call Delete_EmptyRow
        //                    //                    End If
        //                    ////취소
        //                }
        //                else if (oLast_Mode == 101)
        //                {
        //                    //                    If Cancel_oPurchaseOrders(2) = False Then
        //                    //                        oLast_Mode = 0
        //                    //                        BubbleEvent = False
        //                    //                        Exit Sub
        //                    //                    Else
        //                    //                        oLast_Mode = 0
        //                    //                    End If
        //                    ////닫기
        //                }
        //                else if (oLast_Mode == 102)
        //                {
        //                    //                    If Close_oPurchaseOrders(2) = False Then
        //                    //                        oLast_Mode = 0
        //                    //                        BubbleEvent = False
        //                    //                        Exit Sub
        //                    //                    Else
        //                    //                        oLast_Mode = 0
        //                    //                    End If
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                ////36
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((BusinessObjectInfo.BeforeAction == false))
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                ////33
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                ////34
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                ////35
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                ////36
        //                break;
        //        }
        //    }
        //    return;
        //Raise_FormDataEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if ((eventInfo.BeforeAction == true))
        //    {
        //        ////작업
        //    }
        //    else if ((eventInfo.BeforeAction == false))
        //    {
        //        ////작업
        //    }
        //    return;
        //Raise_RightClickEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion
    }
}

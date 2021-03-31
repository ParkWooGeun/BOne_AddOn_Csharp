using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 금속분말 재작업등록
    /// </summary>
    internal class PS_PP047 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Grid oGrid01;
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP047.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP047_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP047");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                //oBaseForm01 = oForm02;
                //oBaseItemUID01 = oItemUID02;
                //oBaseColUID01 = oColUID02;
                //oBaseColRow01 = oColRow02;
                //oBaseTradeType01 = oTradeType02;
                //oBaseItmBsort01 = oItmBsort02;

                PS_PP047_CreateItems();
                PS_PP047_ComboBox_Setting();
                PS_PP047_FormItemEnabled();
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
        private void PS_PP047_CreateItems()
        {
            try
            {
                oGrid01 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PS_PP047");
                oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_PP047");

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

                //제품대분류
                oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");

                //제품중분류
                oForm.DataSources.UserDataSources.Add("ItmMsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmMsort").Specific.DataBind.SetBound(true, "", "ItmMsort");

                //일자
                oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
                oForm.DataSources.UserDataSources.Item("DocDate").Value = DateTime.Now.ToString("yyyyMMdd");

                //사번
                oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

                //사번명
                oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

                //제품코드
                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                //제품코드명
                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

                //불량코드
                oForm.DataSources.UserDataSources.Add("FailCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FailCode").Specific.DataBind.SetBound(true, "", "FailCode");

                //불량명
                oForm.DataSources.UserDataSources.Add("FailName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("FailName").Specific.DataBind.SetBound(true, "", "FailName");

                //공정코드
                oForm.DataSources.UserDataSources.Add("CpCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CpCode").Specific.DataBind.SetBound(true, "", "CpCode");

                //순번
                oForm.DataSources.UserDataSources.Add("Seqno", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10);
                oForm.Items.Item("Seqno").Specific.DataBind.SetBound(true, "", "Seqno");

                //재작업중량
                oForm.DataSources.UserDataSources.Add("Weight", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("Weight").Specific.DataBind.SetBound(true, "", "Weight");

                //비고
                oForm.DataSources.UserDataSources.Add("Remark", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("Remark").Specific.DataBind.SetBound(true, "", "Remark");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP047_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                //사업장콤보박스세팅
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                //품목대분류
                oForm.Items.Item("ItmBsort").Specific.ValidValues.Add("", "");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] Where code ='111' ORDER BY Code", "", false, false);
                oForm.Items.Item("ItmBsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_PP047_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CntcCode").Specific.VALUE = dataHelpClass.User_MSTCOD();
                    oForm.Items.Item("Seqno").Specific.VALUE = "0";
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    ////각모드에따른 아이템설정
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PS_PP047_MTX01
        /// </summary>
        private void PS_PP047_MTX01()
        {
            string sQry;
            string BPLID;
            string DocDate;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                BPLID = oForm.Items.Item("BPLId").Specific.VALUE.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim();
                oGrid01.DataTable.Clear();

                sQry = " EXEC [PS_PP047_01]  '" + BPLID + "', '" + codeHelpClass.Left(DocDate, 6) + "'";

                oGrid01.DataTable.ExecuteQuery(sQry);
                oGrid01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
            }
        }

        /// <summary>
        /// PS_PP047_MTX01
        /// </summary>
        private void PS_PP047_MTX02(string oUID, int oRow = 0, string oCol = "")
        {
            string sQry;
            int sRow;
            int errCode = 0;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                sRow = oRow;
                Param01 = oGrid01.DataTable.Columns.Item("사업장").Cells.Item(oRow).Value;
                Param02 = oGrid01.DataTable.Columns.Item("일자").Cells.Item(oRow).Value;
                Param03 = oGrid01.DataTable.Columns.Item("제품코드").Cells.Item(oRow).Value;
                Param04 = oGrid01.DataTable.Columns.Item("공정코드").Cells.Item(oRow).Value;
                Param05 = Convert.ToString(oGrid01.DataTable.Columns.Item("순번").Cells.Item(oRow).Value);

                sQry = "EXEC PS_PP047_02 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "'";

                oRecordSet01.DoQuery(sQry);

                if (oRecordSet01.RecordCount == 0)
                {
                    PS_PP047_Form_ini();
                    errCode = 1;
                    throw new Exception();
                }
                else
                {
                    oForm.Items.Item("BPLId").Specific.Select(oRecordSet01.Fields.Item("BPLId").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("ItmBsort").Specific.Select(oRecordSet01.Fields.Item("ItmBsort").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("ItmMsort").Specific.Select(oRecordSet01.Fields.Item("ItmMsort").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.DataSources.UserDataSources.Item("DocDate").Value = oRecordSet01.Fields.Item("DocDate").Value;
                    oForm.DataSources.UserDataSources.Item("Seqno").Value = Convert.ToString(oRecordSet01.Fields.Item("Seqno").Value);
                    oForm.DataSources.UserDataSources.Item("CntcCode").Value = oRecordSet01.Fields.Item("CntcCode").Value;
                    oForm.DataSources.UserDataSources.Item("CntcName").Value = oRecordSet01.Fields.Item("CntcName").Value;
                    oForm.DataSources.UserDataSources.Item("Weight").Value =Convert.ToString(oRecordSet01.Fields.Item("Weight").Value);
                    oForm.DataSources.UserDataSources.Item("FailCode").Value = oRecordSet01.Fields.Item("FailCode").Value;
                    oForm.DataSources.UserDataSources.Item("FailName").Value = oRecordSet01.Fields.Item("FailName").Value;
                    oForm.DataSources.UserDataSources.Item("Remark").Value = oRecordSet01.Fields.Item("Remark").Value;
                    oForm.Items.Item("ItemCode").Specific.VALUE = oRecordSet01.Fields.Item("ItemCode").Value;
                    oForm.Items.Item("ItemName").Specific.VALUE = oRecordSet01.Fields.Item("ItemName").Value;
                    oForm.Items.Item("CpCode").Specific.Select(oRecordSet01.Fields.Item("CpCode").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                oForm.ActiveItem = "CntcCode";
                oForm.Items.Item("BPLId").Enabled = false;
                oForm.Items.Item("DocDate").Enabled = false;
                oForm.Items.Item("ItemCode").Enabled = false;
                oForm.Items.Item("CpCode").Enabled = false;

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            }
            catch (Exception ex)
            {
                if (errCode == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
            }
        }


        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP047_SAVE()
        {
            string sQry;
            int Seqno;
            string FailName ;
            string CntcName ;
            string DocDate;
            string ItemName;
            string ItmMsort;
            string BPLID;
            string ItmBsort;
            string ItemCode;
            string CpCode;
            string CntcCode;
            string FailCode;
            string Remark;
            double Weight;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                BPLID = oForm.Items.Item("BPLId").Specific.VALUE.ToString().Trim();
                ItmBsort = oForm.Items.Item("ItmBsort").Specific.VALUE.ToString().Trim();
                ItmMsort = oForm.Items.Item("ItmMsort").Specific.VALUE;
                ItemCode = oForm.Items.Item("ItemCode").Specific.VALUE.ToString().Trim();
                ItemName = oForm.Items.Item("ItemName").Specific.VALUE;
                CpCode = oForm.Items.Item("CpCode").Specific.VALUE.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim();
                Seqno =  Convert.ToInt32(oForm.Items.Item("Seqno").Specific.VALUE);
                CntcCode = oForm.Items.Item("CntcCode").Specific.VALUE.ToString().Trim();
                CntcName = oForm.Items.Item("CntcName").Specific.VALUE;
                FailCode = oForm.Items.Item("FailCode").Specific.VALUE;
                FailName = oForm.Items.Item("FailName").Specific.VALUE;
                Weight = Convert.ToDouble(oForm.Items.Item("Weight").Specific.VALUE);
                Remark = oForm.Items.Item("Remark").Specific.VALUE;

                if (string.IsNullOrEmpty(DocDate.ToString().Trim()))
                {
                    errMessage = "일자가 없습니다. 확인바랍니다";
                    oForm.ActiveItem = "DocDate";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(ItemCode.ToString().Trim()))
                {
                    errMessage = "제품코드가 없습니다. 확인바랍니다.";
                    oForm.ActiveItem = "ItemCode";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(FailCode.ToString().Trim()))
                {
                    errMessage = "재작업코드가 없습니다. 확인바랍니다.";
                    oForm.ActiveItem = "FailCode";
                    throw new Exception();
                }
                if (Weight == 0)
                {
                    errMessage = "중량이 없습니다. 확인바랍니다.";
                    oForm.ActiveItem = "Weight";
                    throw new Exception();
                }

                sQry = " Select Count(*) From [Z_PS_PP047] Where BPLId = '" + BPLID + "' and DocDate = '" + DocDate + "' And ItemCode = '" + ItemCode + "' And CpCode = '" + CpCode + "' And Seqno = " + Seqno + "";
                oRecordSet01.DoQuery(sQry);
                if (oRecordSet01.Fields.Item(0).Value <= 0)
                {
                    sQry = " Select Isnull(Max(Seqno),0) From [Z_PS_PP047] Where BPLId = '" + BPLID + "' And DocDate = '" + DocDate + "' And ItemCode = '" + ItemCode + "' And CpCode = '" + CpCode + "'";
                    oRecordSet01.DoQuery(sQry);

                    Seqno = oRecordSet01.Fields.Item(0).Value + 1;

                    sQry = "INSERT INTO [Z_PS_PP047]";
                    sQry = sQry + " (";
                    sQry = sQry + "BPLId,";
                    sQry = sQry + "ItmBsort,";
                    sQry = sQry + "ItmMsort,";
                    sQry = sQry + "ItemCode,";
                    sQry = sQry + "ItemName,";
                    sQry = sQry + "CpCode,";
                    sQry = sQry + "DocDate,";
                    sQry = sQry + "Seqno,";
                    sQry = sQry + "CntcCode,";
                    sQry = sQry + "CntcName,";
                    sQry = sQry + "Weight,";
                    sQry = sQry + "FailCode,";
                    sQry = sQry + "FailName,";
                    sQry = sQry + "Remark";
                    sQry = sQry + " ) ";
                    sQry = sQry + "VALUES(";

                    sQry = sQry + "'" + BPLID + "',";
                    sQry = sQry + "'" + ItmBsort + "',";
                    sQry = sQry + "'" + ItmMsort + "',";
                    sQry = sQry + "'" + ItemCode + "',";
                    sQry = sQry + "'" + ItemName + "',";
                    sQry = sQry + "'" + CpCode + "',";
                    sQry = sQry + "'" + DocDate + "',";
                    sQry = sQry + Seqno + ",";
                    sQry = sQry + "'" + CntcCode + "',";
                    sQry = sQry + "'" + CntcName + "',";
                    sQry = sQry + Weight + ",";
                    sQry = sQry + "'" + FailCode + "',";
                    sQry = sQry + "'" + FailName + "',";
                    sQry = sQry + "'" + Remark + "'";
                    sQry = sQry + ") ";

                    oRecordSet01.DoQuery(sQry);
                }
                else
                {
                    sQry = "Update [Z_PACKING_PD] set ";
                    sQry = sQry + "CntcCode = '" + CntcCode + "',";
                    sQry = sQry + "CntcName = '" + CntcName + "',";
                    sQry = sQry + "FailCode = '" + FailCode + "',";
                    sQry = sQry + "FailName = '" + FailName + "',";
                    sQry = sQry + "Remark = '" + Remark + "'";

                    sQry = sQry + " Where BPLId = '" + BPLID + "' and DocDate = '" + DocDate + "' And ItemCode = '" + ItemCode + "' And CpCode = '" + CpCode + "' And Seqno = " + Seqno + "";
                    oRecordSet01.DoQuery(sQry);
                }
                PS_PP047_FormItemEnabled();
            }
            catch (Exception ex)
            {
                if (errMessage != null)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP047_Delete()
        {
            string CpCode;
            string DocDate;
            string BPLID ;
            string ItemCode;
            int Seqno;
            int Cnt;
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                BPLID = oForm.Items.Item("BPLId").Specific.VALUE.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim();
                ItemCode = oForm.Items.Item("ItemCode").Specific.VALUE.ToString().Trim();
                CpCode = oForm.Items.Item("CpCode").Specific.VALUE.ToString().Trim();
                Seqno = Convert.ToInt32(oForm.Items.Item("Seqno").Specific.VALUE);

                sQry = " Select Count(*) From [Z_PS_PP047] Where BPLId = '" + BPLID + "' And DocDate = '" + DocDate + "' And ItemCode = '" + ItemCode + "' And CpCode = '" + CpCode + "' And Seqno = " + Seqno + "";
                oRecordSet01.DoQuery(sQry);

                Cnt = Convert.ToInt32(oRecordSet01.Fields.Item(0).Value);
                if (Cnt > 0)
                {
                    if (PSH_Globals.SBO_Application.MessageBox(" 선택한라인을 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1"))
                    {
                        sQry = "Delete From [Z_PS_PP047] Where BPLId = '" + BPLID + "' And DocDate = '" + DocDate + "' And ItemCode = '" + ItemCode + "' And CpCode = '" + CpCode + "' And Seqno = " + Seqno + "";
                        oRecordSet01.DoQuery(sQry);
                    }
                }
                else
                {
                    errMessage = "조회후 삭제하십시요.";
                    throw new Exception();
                }
                PS_PP047_MTX01();
                PS_PP047_Form_ini();
            }
            catch (Exception ex)
            {
                if(errMessage != null)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP047_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP047'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP047_Form_ini()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Items.Item("CntcCode").Specific.VALUE = dataHelpClass.User_MSTCOD();

                oForm.DataSources.UserDataSources.Item("ItemCode").Value = "";
                oForm.DataSources.UserDataSources.Item("ItemName").Value = "";
                oForm.DataSources.UserDataSources.Item("Seqno").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("FailCode").Value = "";
                oForm.DataSources.UserDataSources.Item("FailName").Value = "";

                oForm.DataSources.UserDataSources.Item("Weight").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("Remark").Value = "";

                oForm.DataSources.UserDataSources.Item("CpCode").Value = "";
                oForm.DataSources.UserDataSources.Item("ItmBsort").Value = "";
                oForm.DataSources.UserDataSources.Item("ItmMsort").Value = "";

                //Key enable
                oForm.Items.Item("BPLId").Enabled = true;
                oForm.Items.Item("DocDate").Enabled = true;
                oForm.Items.Item("ItemCode").Enabled = true;
                oForm.Items.Item("CpCode").Enabled = true;

                oForm.ActiveItem = "DocDate";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
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
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                //    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                    if (pVal.ItemUID == "Btn_ret")
                    {
                        PS_PP047_MTX01();
                    }
                    if (pVal.ItemUID == "Btn01")
                    {
                        PS_PP047_SAVE();
                        PS_PP047_MTX01();
                    }
                    if (pVal.ItemUID == "Btn_del")
                    {
                        PS_PP047_Delete();
                        PS_PP047_FormItemEnabled();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PS_PP047")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
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
                        //사번
                        if (pVal.ItemUID == "CntcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.VALUE))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
                                BubbleEvent = false;
                            }
                        }
                        //제품코드
                        if (pVal.ItemUID == "ItemCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.VALUE))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
                                BubbleEvent = false;
                            }
                        }
                        //재작업사유(불량)코드
                        if (pVal.ItemUID == "FailCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("FailCode").Specific.VALUE))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
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
                if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02")
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "ItmBsort")
                    {
                        for (i = oForm.Items.Item("ItmMsort").Specific.ValidValues.Count - 1; i>= 0; i += -1)
                        {
                            oForm.Items.Item("ItmMsort").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("ItmMsort").Specific.ValidValues.Add("", "");
                        dataHelpClass.Set_ComboList(oForm.Items.Item("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" + oForm.Items.Item("ItmBsort").Specific.Selected.VALUE + "' ORDER BY U_Code", "", false, false);
                        if (oForm.Items.Item("ItmMsort").Specific.ValidValues.Count > 0)
                        {
                            oForm.Items.Item("ItmMsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.Row > 0)
                            {
                                //     PS_PP047_MTX02 pVal.ItemUID, pVal.Row, pVal.ColUID
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            PS_PP047_MTX02(pVal.ItemUID, pVal.Row, pVal.ColUID);
                        }

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
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
            int i;
            string oQuery;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
                        //사원명
                        if (pVal.ItemUID == "CntcCode")
                        {
                            oQuery = "SELECT FullName = U_FullName  ";
                            oQuery = oQuery + "FROM [@PH_PY001A] WHERE Code = '" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'";
                            oRecordSet01.DoQuery(oQuery);
                            oForm.Items.Item("CntcName").Specific.VALUE = oRecordSet01.Fields.Item("FullName").Value.ToString().Trim();
                        }
                        //제품코드
                        if (pVal.ItemUID == "ItemCode")
                        {
                            oQuery = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'";
                            oRecordSet01.DoQuery(oQuery);
                            oForm.Items.Item("ItemName").Specific.VALUE = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                            for (i = oForm.Items.Item("CpCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                            {
                                oForm.Items.Item("CpCode").Specific.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                            oForm.Items.Item("CpCode").Specific.ValidValues.Add("", "");
                            dataHelpClass.Set_ComboList(oForm.Items.Item("CpCode").Specific, "SELECT U_CpCode, U_CpName FROM [@PS_PP004H] WHERE U_ItemCode = '" + oForm.Items.Item("ItemCode").Specific.VALUE.ToString().Trim() + "' ORDER BY U_Sequence", "", false, false);
                            if (oForm.Items.Item("CpCode").Specific.ValidValues.Count > 0)
                            {
                                oForm.Items.Item("CpCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }
                        //불량코드
                        if (pVal.ItemUID == "FailCode")
                        {
                            oQuery = "Select U_SmalName From [@PS_PP003L] Where U_SmalCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'";
                            oRecordSet01.DoQuery(oQuery);
                            oForm.Items.Item("FailName").Specific.String = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    PS_PP047_FormItemEnabled();
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
                    SubMain.Remove_Forms(oFormUniqueID);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                                     //추가
                            oGrid01.DataTable.Clear();
                            PS_PP047_Form_ini();
                            oForm.Items.Item("Btn_ret").Click();
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
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
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                        case "1287": //복제
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
                if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02")
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
    }
}

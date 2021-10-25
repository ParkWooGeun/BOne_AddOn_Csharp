using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using System.Collections.Generic;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    ///분말포장지시등록
    /// </summary>
    internal class PS_PACKING_PD : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Grid oGrid01;

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PACKING_PD.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PACKING_PD_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PACKING_PD");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_PACKING_PD_CreateItems();
                PS_PACKING_PD_ComboBox_Setting();
                PS_PACKING_PD_FormItemEnabled();

                oForm.EnableMenu("1282", true);
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
        private void PS_PACKING_PD_CreateItems()
        {
            try
            {
                oGrid01 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("ZTEMP1");

                oGrid01.DataTable = oForm.DataSources.DataTables.Item("ZTEMP1");

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

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

                //작업지시번호
                oForm.DataSources.UserDataSources.Add("OrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("OrdNum").Specific.DataBind.SetBound(true, "", "OrdNum");

                //작지번호
                oForm.DataSources.UserDataSources.Add("PP030HNo", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10);
                oForm.Items.Item("PP030HNo").Specific.DataBind.SetBound(true, "", "PP030HNo");

                //작지순번
                oForm.DataSources.UserDataSources.Add("PP030MNo", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10);
                oForm.Items.Item("PP030MNo").Specific.DataBind.SetBound(true, "", "PP030MNo");

                //포장단위
                oForm.DataSources.UserDataSources.Add("BoxDiv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("BoxDiv").Specific.DataBind.SetBound(true, "", "BoxDiv");

                //포장단위Kg
                oForm.DataSources.UserDataSources.Add("BoxKg", SAPbouiCOM.BoDataType.dt_QUANTITY, 20);
                oForm.Items.Item("BoxKg").Specific.DataBind.SetBound(true, "", "BoxKg");

                //총포장중량
                oForm.DataSources.UserDataSources.Add("Quantity", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("Quantity").Specific.DataBind.SetBound(true, "", "Quantity");

                //박스수량
                oForm.DataSources.UserDataSources.Add("BoxCnt", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10);
                oForm.Items.Item("BoxCnt").Specific.DataBind.SetBound(true, "", "BoxCnt");

                //배치번호
                oForm.DataSources.UserDataSources.Add("BatchNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("BatchNum").Specific.DataBind.SetBound(true, "", "BatchNum");

                //이전배치번호
                oForm.DataSources.UserDataSources.Add("bBatchNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("bBatchNum").Specific.DataBind.SetBound(true, "", "bBatchNum");

                //검사의뢰번호
                oForm.DataSources.UserDataSources.Add("InspNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("InspNo").Specific.DataBind.SetBound(true, "", "InspNo");

                //거래처코드
                oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

                //거래처명
                oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

                //포장단위
                oForm.DataSources.UserDataSources.Add("SaleType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SaleType").Specific.DataBind.SetBound(true, "", "SaleType");

                //고객순번
                oForm.DataSources.UserDataSources.Add("CardSeq", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("CardSeq").Specific.DataBind.SetBound(true, "", "CardSeq");

                oForm.Items.Item("CardSeq").Specific.Value = "00";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PACKING_PD_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.Items.Item("SaleType").Specific.ValidValues.Add("", "");
                dataHelpClass.Set_ComboList(oForm.Items.Item("SaleType").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001L] b Where b.Code = 'P212' and b.U_UseYN = 'Y' Order by U_Seq", "", false, false);

                oForm.Items.Item("BoxDiv").Specific.ValidValues.Add("", "");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BoxDiv").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001L] b Where b.Code = 'P204' and b.U_UseYN = 'Y' Order by U_Seq", "", false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_PACKING_PD_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                    oForm.Items.Item("BoxKg").Enabled = false;
                    oForm.Items.Item("BatchNum").Enabled = false;
                    oForm.Items.Item("CardSeq").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("BoxKg").Enabled = false;
                    oForm.Items.Item("BatchNum").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("BoxKg").Enabled = false;
                    oForm.Items.Item("BatchNum").Enabled = false;
                }

                if (oForm.Items.Item("SaleType").Specific.Value.ToString().Trim() == "1")
                {
                    oForm.Items.Item("bBatchNum").Enabled = false;
                }
                else
                {
                    oForm.Items.Item("bBatchNum").Enabled = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_PACKING_PD_MTX01
        /// </summary>
        private void PS_PACKING_PD_MTX01()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                oGrid01.DataTable.Clear();
                sQry = " EXEC [PS_PACKING_PD_01]  '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "', '" + oForm.Items.Item("DocDate").Specific.Value.ToString().Trim() + "'";
                
                oGrid01.DataTable.ExecuteQuery(sQry);
                oGrid01.AutoResizeColumns();

                oGrid01.Columns.Item(9).RightJustified = true;
                oGrid01.Columns.Item(10).RightJustified = true;
                oGrid01.Columns.Item(11).RightJustified = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
            }
        }

        /// <summary>
        /// PS_PACKING_PD_MTX01
        /// </summary>
        private void PS_PACKING_PD_MTX02(string oUID, int oRow, string oCol)
        {
            string errMessage = string.Empty;
            int sRow;
            string sQry;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            string Param06;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                sRow = oRow;
                PS_PACKING_PD_Form_ini();
                Param01 = oGrid01.DataTable.Columns.Item("사업장").Cells.Item(oRow).Value;
                Param02 = oGrid01.DataTable.Columns.Item("일자").Cells.Item(oRow).Value;
                Param03 = oGrid01.DataTable.Columns.Item("제품코드").Cells.Item(oRow).Value;
                Param04 = oGrid01.DataTable.Columns.Item("제품명").Cells.Item(oRow).Value;
                Param05 = oGrid01.DataTable.Columns.Item("작업지시번호").Cells.Item(oRow).Value;
                Param06 = oGrid01.DataTable.Columns.Item("배치번호").Cells.Item(oRow).Value;

                sQry = "EXEC PS_PACKING_PD_02 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "'";

                oRecordSet01.DoQuery(sQry);

                if (oRecordSet01.RecordCount == 0)
                {
                    PS_PACKING_PD_Form_ini();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                else
                {
                    oForm.Items.Item("BPLId").Specific.Select(oRecordSet01.Fields.Item("BPLId").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.DataSources.UserDataSources.Item("DocDate").Value = oRecordSet01.Fields.Item("DocDate").Value;
                    oForm.DataSources.UserDataSources.Item("ItemCode").Value = oRecordSet01.Fields.Item("ItemCode").Value;
                    oForm.DataSources.UserDataSources.Item("ItemName").Value = oRecordSet01.Fields.Item("ItemName").Value;
                    oForm.DataSources.UserDataSources.Item("OrdNum").Value = oRecordSet01.Fields.Item("OrdNum").Value;
                    oForm.DataSources.UserDataSources.Item("BatchNUm").Value = oRecordSet01.Fields.Item("BatchNum").Value;
                    oForm.DataSources.UserDataSources.Item("bBatchNUm").Value = oRecordSet01.Fields.Item("bBatchNum").Value;
                    oForm.DataSources.UserDataSources.Item("PP030HNo").Value = oRecordSet01.Fields.Item("PP030HNo").Value.ToString().Trim();
                    oForm.DataSources.UserDataSources.Item("PP030MNo").Value = oRecordSet01.Fields.Item("PP030MNo").Value.ToString().Trim();
                    oForm.DataSources.UserDataSources.Item("InspNo").Value = oRecordSet01.Fields.Item("InspNo").Value;
                    oForm.DataSources.UserDataSources.Item("BoxDiv").Value = oRecordSet01.Fields.Item("BoxDiv").Value;
                    oForm.DataSources.UserDataSources.Item("BoxKg").Value = oRecordSet01.Fields.Item("BoxKg").Value.ToString().Trim();
                    oForm.DataSources.UserDataSources.Item("Quantity").Value = oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim();
                    oForm.DataSources.UserDataSources.Item("BoxCnt").Value = oRecordSet01.Fields.Item("BoxCnt").Value.ToString().Trim();
                    oForm.DataSources.UserDataSources.Item("CntcCode").Value = oRecordSet01.Fields.Item("CntcCode").Value;
                    oForm.DataSources.UserDataSources.Item("CntcName").Value = oRecordSet01.Fields.Item("CntcName").Value;
                    oForm.DataSources.UserDataSources.Item("CardCode").Value = oRecordSet01.Fields.Item("CardCode").Value;
                    oForm.DataSources.UserDataSources.Item("CardName").Value = oRecordSet01.Fields.Item("CardName").Value;
                    oForm.DataSources.UserDataSources.Item("CardSeq").Value = oRecordSet01.Fields.Item("CardSeq").Value;
                    oForm.DataSources.UserDataSources.Item("SaleType").Value = oRecordSet01.Fields.Item("SaleType").Value;
                }
                oForm.ActiveItem = "CntcCode";
                oForm.Items.Item("BPLId").Enabled = false;
                oForm.Items.Item("DocDate").Enabled = false;
                oForm.Items.Item("ItemCode").Enabled = false;
                oForm.Items.Item("OrdNum").Enabled = false;
                oForm.Items.Item("InspNo").Enabled = false;

                if (oForm.Items.Item("SaleType").Specific.Value.ToString().Trim() == "1")
                {
                    oForm.Items.Item("bBatchNum").Enabled = false;
                }
                else
                {
                    oForm.Items.Item("bBatchNum").Enabled = true;
                }

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PACKING_PD_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PACKING_PD'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private void PS_PACKING_PD_Form_ini()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                oForm.DataSources.UserDataSources.Item("ItemCode").Value = "";
                oForm.DataSources.UserDataSources.Item("ItemName").Value = "";
                oForm.DataSources.UserDataSources.Item("OrdNum").Value = "";
                oForm.DataSources.UserDataSources.Item("PP030HNo").Value = "";
                oForm.DataSources.UserDataSources.Item("PP030MNo").Value = "";
                oForm.DataSources.UserDataSources.Item("InspNo").Value = "";

                oForm.DataSources.UserDataSources.Item("BoxKg").Value = "0";
                oForm.DataSources.UserDataSources.Item("Quantity").Value = "0";
                oForm.DataSources.UserDataSources.Item("BoxCnt").Value = "0";
                oForm.DataSources.UserDataSources.Item("BatchNum").Value = "";
                oForm.DataSources.UserDataSources.Item("bBatchNum").Value = "";
                oForm.Items.Item("BoxDiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.Items.Item("BPLId").Enabled = true;
                oForm.Items.Item("DocDate").Enabled = true;
                oForm.Items.Item("ItemCode").Enabled = true;
                oForm.Items.Item("OrdNum").Enabled = true;
                oForm.Items.Item("InspNo").Enabled = true;
                oForm.ActiveItem = "DocDate";
                oForm.Update();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_PACKING_PD_InspNo
        /// </summary>
        /// <returns></returns>
        private void PS_PACKING_PD_InspNo()
        {
            int Seq;
            string sQry;
            string BPLID;
            string DocDate;
            string InspNo;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

                sQry = " Select Isnull(right(Max(InspNo),2),'') From [Z_PACKING_PD] Where BPLId = '" + BPLID + "' and DocDate = '" + DocDate + "'";

                oRecordSet01.DoQuery(sQry);
                if (string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value))
                {
                    Seq = 1;
                }
                else
                {
                    Seq = Convert.ToInt32(oRecordSet01.Fields.Item(0).Value) + 1;
                }
                InspNo = DocDate + Seq.ToString().PadLeft(2,'0');

                oForm.Items.Item("InspNo").Specific.Value = InspNo;
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
        /// PS_PACKING_PD_Delete
        /// </summary>
        /// <returns></returns>
        private void PS_PACKING_PD_Delete()
        {
            string errMessage = string.Empty;
            string OrdNum;
            string ItemCode;
            string BPLID;
            string DocDate;
            string ItemName;
            string BatchNum;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                ItemName = oForm.Items.Item("ItemName").Specific.Value;
                OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
                BatchNum = oForm.Items.Item("BatchNum").Specific.Value;

                sQry = "select * from [@PS_QM008H] a where a.Status ='O' and a.Canceled ='N' and a.U_InspNo ='" + oForm.Items.Item("InspNo").Specific.Value + "'";
                oRecordSet01.DoQuery(sQry);

                if (Convert.ToInt16(oRecordSet01.Fields.Item(0).Value) > 0)
                {
                    errMessage = "이미 검사성적서가 등록되어있습니다. 등록,수정시 검사성적서를 취소 후 수정하세요.";
                    throw new Exception();
                }

                sQry = " Select Count(*) From [Z_PACKING_PD] Where BPLId = '" + BPLID + "' And DocDate = '" + DocDate + "' And ItemCode = '" + ItemCode + "' And ItemName = '" + ItemName + "' And OrdNum = '" + OrdNum + "' And BatchNum = '" + BatchNum + "'";
                oRecordSet01.DoQuery(sQry);

                if (Convert.ToInt16(oRecordSet01.Fields.Item(0).Value) > 0)
                {
                    if (PSH_Globals.SBO_Application.MessageBox("선택한 라인을 삭제하시겠습니까?", 2, "예", "아니오") == 1)
                    {
                        sQry = "Delete From [Z_PACKING_PD] Where BPLId = '" + BPLID + "' And DocDate = '" + DocDate + "' And ItemCode = '" + ItemCode + "' And ItemName = '" + ItemName + "' And OrdNum = '" + OrdNum + "' And BatchNum = '" + BatchNum + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }
                else
                {
                    errMessage = "조회후 삭제하십시요.";
                    throw new Exception();
                }
                PS_PACKING_PD_MTX01();
                PS_PACKING_PD_Form_ini();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_PACKING_PD_SAVE
        /// </summary>
        private void PS_PACKING_PD_SAVE()
        {
            string errMessage = string.Empty;
            int PP030HNo;
            int PP030MNo;
            int BoxCnt;
            string InspNo;
            string BoxDiv;
            string ItemName;
            string CntcName;
            string DocDate;
            string BPLID;
            string CntcCode;
            string ItemCode;
            string OrdNum;
            string BatchNum;
            string bBatchNum;
            string CardSeq;
            string CardName;
            string CardCode;
            string SaleType;
            string PP080YN;
            string sQry;
            double Boxkg;
            double Quantity;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
                CntcName = oForm.Items.Item("CntcName").Specific.Value;
                ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                ItemName = oForm.Items.Item("ItemName").Specific.Value;
                OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
                PP030HNo = Convert.ToInt32(oForm.Items.Item("PP030HNo").Specific.Value);
                PP030MNo = Convert.ToInt32(oForm.Items.Item("PP030MNo").Specific.Value);
                InspNo = oForm.Items.Item("InspNo").Specific.Value.ToString().Trim();
                BoxDiv = oForm.Items.Item("BoxDiv").Specific.Value;
                Boxkg = Convert.ToDouble(oForm.Items.Item("BoxKg").Specific.Value);
                Quantity = Convert.ToDouble(oForm.Items.Item("Quantity").Specific.Value);
                BoxCnt = Convert.ToInt32(oForm.Items.Item("BoxCnt").Specific.Value);
                BatchNum = oForm.Items.Item("BatchNum").Specific.Value;
                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                CardName = oForm.Items.Item("CardName").Specific.Value;
                bBatchNum = oForm.Items.Item("bBatchNum").Specific.Value;
                CardSeq = oForm.Items.Item("CardSeq").Specific.Value;
                SaleType = oForm.Items.Item("SaleType").Specific.Value;

                sQry = "select * from [@PS_QM008H] a where a.Status ='O' and a.Canceled ='N' and a.U_InspNo ='" + oForm.Items.Item("InspNo").Specific.Value + "'";
                oRecordSet01.DoQuery(sQry);

                if (oRecordSet01.RecordCount > 0)
                {
                    errMessage = "이미 검사성적서가 등록되어있습니다. 등록,수정시 검사성적서를 취소 후 수정하세요.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(DocDate.ToString().Trim()))
                {
                    errMessage = "일자가 없습니다. 확인바랍니다.";
                    oForm.ActiveItem = "DocDate";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(OrdNum.ToString().Trim()))
                {
                    errMessage = "작업지시번호가 없습니다. 확인바랍니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(InspNo.ToString().Trim()))
                {
                    errMessage = "검사의뢰번호가 없습니다. 확인바랍니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(SaleType.ToString().Trim()))
                {
                    errMessage = "판매유형이 없습니다. 확인바랍니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(Quantity.ToString().Trim()))
                {
                    errMessage = "포장중량이 없습니다. 확인바랍니다.";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(CardSeq.ToString().Trim()))
                {
                    errMessage = "고객순번이 없습니다. 확인바랍니다.";
                    throw new Exception();
                }
                if (!string.IsNullOrEmpty(bBatchNum.ToString().Trim()))
                {
                    PP080YN = "Y";
                }
                else
                {
                    PP080YN = "N";
                }

                sQry = " Select Count(*) From [Z_PACKING_PD] Where BPLId = '" + BPLID + "' and DocDate = '" + DocDate + "' And ItemCode = '" + ItemCode + "' And ItemName = '" + ItemName + "' And OrdNum = '" + OrdNum + "' And BatchNum = '" + BatchNum + "'";
                oRecordSet01.DoQuery(sQry);
                if (Convert.ToInt16(oRecordSet01.Fields.Item(0).Value) <= 0)
                {
                    sQry = "INSERT INTO [Z_PACKING_PD]";
                    sQry += " (";
                    sQry += "BPLId,";
                    sQry += "DocDate,";
                    sQry += "CntcCode,";
                    sQry += "CntcName,";
                    sQry += "ItemCode,";
                    sQry += "ItemName,";
                    sQry += "OrdNum,";
                    sQry += "PP030HNo,";
                    sQry += "PP030MNo,";
                    sQry += "InspNo,";
                    sQry += "BoxDiv,";
                    sQry += "BoxKg,";
                    sQry += "Quantity,";
                    sQry += "BoxCnt,";
                    sQry += "BatchNum,";
                    sQry += "CardCode,";
                    sQry += "CardName,";
                    sQry += "CardSeq,";
                    sQry += "PP080YN,";
                    sQry += "bBatchNum,";
                    sQry += "SaleType";
                    sQry += " ) ";
                    sQry += "VALUES(";

                    sQry += "'" + BPLID + "',";
                    sQry += "'" + DocDate + "',";
                    sQry += "'" + CntcCode + "',";
                    sQry += "'" + CntcName + "',";
                    sQry += "'" + ItemCode + "',";
                    sQry += "'" + ItemName + "',";
                    sQry += "'" + OrdNum + "',";
                    sQry += PP030HNo + ",";
                    sQry += PP030MNo + ",";
                    sQry += "'" + InspNo + "',";
                    sQry += "'" + BoxDiv + "',";
                    sQry += Boxkg + ",";
                    sQry += Quantity + ",";
                    sQry += BoxCnt + ",";
                    sQry += "'" + BatchNum + "',";
                    sQry += "'" + CardCode + "',";
                    sQry += "'" + CardName + "',";
                    sQry += "'" + CardSeq + "',";
                    sQry += "'" + PP080YN + "',";
                    sQry += "'" + bBatchNum + "',";
                    sQry += "'" + SaleType + "'";
                    sQry += ") ";

                    oRecordSet01.DoQuery(sQry);
                }
                else
                {
                    //제품검사가 된경우 저장안되게 해야함
                    sQry = "Update [Z_PACKING_PD] set ";
                    sQry += "PP030HNo = '" + PP030HNo + "',";
                    sQry += "PP030MNo = '" + PP030MNo + "',";
                    sQry += "InspNo = '" + InspNo + "',";
                    sQry += "BoxDiv = '" + BoxDiv + "',";
                    sQry += "BoxKg = '" + Boxkg + "',";
                    sQry += "Quantity = '" + Quantity + "',";
                    sQry += "CntcCode = '" + CntcCode + "',";
                    sQry += "CntcName = '" + CntcName + "',";
                    sQry += "CardCode = '" + CardCode + "',";
                    sQry += "CardName = '" + CardName + "',";
                    sQry += "CardSeq = '" + CardSeq + "',";
                    sQry += "BoxCnt = '" + BoxCnt + "',";
                    sQry += "SaleType = '" + SaleType + "'";
                    sQry += " Where BPLId = '" + BPLID + "' and DocDate = '" + DocDate + "' And ItemCode = '" + ItemCode + "' And ItemName = '" + ItemName + "' And OrdNum = '" + OrdNum + "' And BatchNum = '" + BatchNum + "'";
                    oRecordSet01.DoQuery(sQry);
                }
                PS_PACKING_PD_FormItemEnabled();
                PS_PACKING_PD_MTX01();
                PS_PACKING_PD_Form_ini();

                oForm.Items.Item("DocDate").Specific.Value = DocDate;
                PS_PACKING_PD_MTX01();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_PACKING_PD_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            string Param06;
            string Param07;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                Param01 = oForm.Items.Item("BPLId").Specific.Selected.Value;
                Param02 = oForm.Items.Item("DocDate").Specific.Value;
                Param03 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                Param04 = oForm.Items.Item("ItemName").Specific.Value;
                Param05 = oForm.Items.Item("OrdNum").Specific.Value;
                Param06 = oForm.Items.Item("BatchNum").Specific.Value;
                Param07 = oForm.Items.Item("CardCode").Specific.Value;

                WinTitle = "BOX-LABEL출력[PS_PACKING_PD_05] ";
                ReportName = "PS_PACKING_PD_05.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter

                dataPackParameter.Add(new PSH_DataPackClass("@BPLId", Param01));
                dataPackParameter.Add(new PSH_DataPackClass("@DocDate", Param02));
                dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", Param03));
                dataPackParameter.Add(new PSH_DataPackClass("@ItemName", Param04));
                dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", Param05));
                dataPackParameter.Add(new PSH_DataPackClass("@BatchNum", Param06));
                dataPackParameter.Add(new PSH_DataPackClass("@CardCode", Param07));

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, "Y");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    if (pVal.ItemUID == "Btn_ret")
                    {
                        PS_PACKING_PD_MTX01();
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        PS_PACKING_PD_SAVE();
                    }
                    else if (pVal.ItemUID == "Btn_del")
                    {
                        PS_PACKING_PD_Delete();
                        PS_PACKING_PD_FormItemEnabled();
                    }
                    else if (pVal.ItemUID == "Btn_Prt")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_PACKING_PD_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    else if (pVal.ItemUID == "Btn_Create")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_PACKING_PD_InspNo();
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.MessageBox("신규모드일때만 가능합니다.");
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                        if (pVal.ItemUID == "CntcCode")//사번
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "ItemCode")  //제품코드
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "CardCode") //거래처코드
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "OrdNum") //작업지시번호
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "InspNo")//검사의뢰번호
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("InspNo").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "CardSeq") //고객순번
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardSeq").Specific.Value))
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string Box;
            string BatchNum;
            string oQuery;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
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
                        if (pVal.ItemUID == "BoxDiv")
                        {
                            if (oForm.Items.Item("BoxDiv").Specific.Value.ToString().Trim() == "S")
                            {
                                oForm.Items.Item("BoxKg").Enabled = true;
                                oForm.Items.Item("BatchNum").Enabled = true;
                            }
                            else
                            {
                                oForm.Items.Item("BoxKg").Enabled = false;
                                oForm.Items.Item("BatchNum").Enabled = false;
                            }

                            oQuery = "Select U_RelCd FROM [@PS_SY001L] b Where b.Code = 'P204' and b.U_Minor = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(oQuery);
                            oForm.Items.Item("BoxKg").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                            if (oForm.Items.Item("Quantity").Specific.Value.ToString().Trim() != "0")
                            {
                                oForm.Items.Item("BoxCnt").Specific.Value = Convert.ToString(Convert.ToDouble(oForm.Items.Item("Quantity").Specific.Value) / Convert.ToDouble(oForm.Items.Item("BoxKg").Specific.Value));
                            }
                            if (!string.IsNullOrEmpty(oForm.Items.Item("InspNo").Specific.Value.ToString().Trim()))
                            {
                                Box = oForm.Items.Item("BoxDiv").Specific.Value.ToString().Trim();
                                BatchNum = Box + codeHelpClass.Right(oForm.Items.Item("InspNo").Specific.Value, 8);
                                oForm.Items.Item("BatchNum").Specific.Value = BatchNum;
                            }
                        }
                        if (pVal.ItemUID == "SaleType")
                        {
                            if (oForm.Items.Item("SaleType").Specific.Value.ToString().Trim() == "1")
                            {
                                oForm.Items.Item("bBatchNum").Enabled = false;
                            }
                            else
                            {
                                oForm.Items.Item("bBatchNum").Enabled = true;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            PS_PACKING_PD_MTX02(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string oQuery;
            string CardCode;
            string ItemCode;
            string Box;
            string Boxkg;
            string BatchNum;
            string DocDate;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CntcCode")//사원명
                        {
                            oQuery = "SELECT FullName = U_FullName  ";
                            oQuery += "FROM [@PH_PY001A] WHERE Code = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                            oRecordSet01.DoQuery(oQuery);
                            oForm.Items.Item("CntcName").Specific.Value = oRecordSet01.Fields.Item("FullName").Value.ToString().Trim();
                        }
                        else if (pVal.ItemUID == "ItemCode") //제품코드
                        {
                            ItemCode = oForm.Items.Item(pVal.ItemUID).Specific.Value;
                            CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();

                            oQuery = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                            oRecordSet01.DoQuery(oQuery);
                            oForm.Items.Item("ItemName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            oForm.Items.Item("OrdNum").Specific.Value = "";
                            oForm.Items.Item("PP030HNo").Specific.Value = "";
                            oForm.Items.Item("PP030MNo").Specific.Value = "";

                            oQuery = "Select Count(*) From [@PS_QM007H] Where U_CardCode = '" + CardCode + "' and U_ItemCode = '" + ItemCode + "'";
                            oRecordSet01.DoQuery(oQuery);
                            if (Convert.ToInt16(oRecordSet01.Fields.Item(0).Value) > 1)
                            {
                                oForm.Items.Item("CardSeq").Enabled = true;
                                oForm.Items.Item("CardSeq").Specific.Value = "";
                            }
                            else
                            {
                                oForm.Items.Item("CardSeq").Enabled = false;
                                oForm.Items.Item("CardSeq").Specific.Value = "00";
                            }
                        }
                        else if (pVal.ItemUID == "CardCode") //거래처코드
                        {
                            CardCode = oForm.Items.Item(pVal.ItemUID).Specific.Value;
                            ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                            oQuery = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                            oRecordSet01.DoQuery(oQuery);
                            oForm.Items.Item("CardName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                            oQuery = "Select Count(*) From [@PS_QM007H] Where U_CardCode = '" + CardCode + "' and U_ItemCode = '" + ItemCode + "'";
                            oRecordSet01.DoQuery(oQuery);
                            if (Convert.ToInt16(oRecordSet01.Fields.Item(0).Value) > 1)
                            {
                                oForm.Items.Item("CardSeq").Enabled = true;
                                oForm.Items.Item("CardSeq").Specific.Value = "";
                            }
                            else
                            {
                                oForm.Items.Item("CardSeq").Enabled = false;
                                oForm.Items.Item("CardSeq").Specific.Value = "00";
                            }
                        }
                        else if (pVal.ItemUID == "Quantity") //총포장중량
                        {
                            if (Convert.ToDouble(oForm.Items.Item(pVal.ItemUID).Specific.Value) > 0)
                            {
                                if (Convert.ToDouble(oForm.Items.Item(pVal.ItemUID).Specific.Value) < Convert.ToDouble(oForm.Items.Item("BoxKg").Specific.Value))
                                {
                                    oForm.Items.Item("BoxCnt").Specific.Value = 1;
                                }
                                else
                                {
                                    if (Convert.ToDouble(oForm.Items.Item(pVal.ItemUID).Specific.Value) % Convert.ToDouble(oForm.Items.Item("BoxKg").Specific.Value) == 0)
                                    {
                                        oForm.Items.Item("BoxCnt").Specific.Value = Convert.ToInt32(Convert.ToDouble(oForm.Items.Item(pVal.ItemUID).Specific.Value) / Convert.ToDouble(oForm.Items.Item("BoxKg").Specific.Value));
                                    }
                                    else
                                    {
                                        oForm.Items.Item("BoxCnt").Specific.Value = Convert.ToInt32(Math.Truncate(Convert.ToDouble(oForm.Items.Item(pVal.ItemUID).Specific.Value) / Convert.ToDouble(oForm.Items.Item("BoxKg").Specific.Value))) + 1;
                                    }
                                }
                            }
                        }
                        else if (pVal.ItemUID == "OrdNum") //작업지시조회
                        {
                            if (!string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                            {
                                oQuery = "Select a.DocEntry, b.LineId From [@PS_PP030H] a Inner Join [@PS_PP030M] b On a.DocEntry = b.DocEntry and a.Canceled = 'N' ";
                                oQuery += " Where a.U_OrdNum = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                                oQuery += " And b.U_CpCode = 'CP80199'";
                                oRecordSet01.DoQuery(oQuery);
                                oForm.Items.Item("PP030HNo").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                                oForm.Items.Item("PP030MNo").Specific.Value = oRecordSet01.Fields.Item(1).Value.ToString().Trim();
                            }
                            else
                            {
                                oForm.Items.Item("PP030HNo").Specific.Value = "";
                                oForm.Items.Item("PP030MNo").Specific.Value = "";
                            }
                        }
                        else if (pVal.ItemUID == "InspNo") //검사의뢰번호
                        {
                            if (!string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                            {
                                Box = oForm.Items.Item("BoxDiv").Specific.Value.ToString().Trim();
                                Boxkg = oForm.Items.Item("BoxKg").Specific.Value.ToString().Trim();
                                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                                oQuery = "Select Isnull(Right(Max(Right(Batchnum,8)),2),'') From Z_PACKING_PD Where DocDate = '" + DocDate + "'";
                                oRecordSet01.DoQuery(oQuery);

                                BatchNum = Box + codeHelpClass.Mid(oForm.Items.Item("InspNo").Specific.Value, 2, 8) + codeHelpClass.Left(Boxkg, Boxkg.IndexOf(".")).PadLeft(4, '0');

                                oForm.Items.Item("BatchNum").Specific.Value = BatchNum;
                            }
                            else
                            {
                                oForm.Items.Item("BatchNum").Specific.Value = "";
                            }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    PS_PACKING_PD_FormItemEnabled();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
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
                            oGrid01.DataTable.Clear();
                            PS_PACKING_PD_Form_ini();
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

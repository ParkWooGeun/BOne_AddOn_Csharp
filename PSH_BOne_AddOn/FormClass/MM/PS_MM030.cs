using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 품의[발주]
    /// </summary>
    internal class PS_MM030 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_MM030H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM030L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private SAPbouiCOM.BoFormMode oLast_Mode;
        private string oDocNum;
        private string LastPick = "";
        private short g_deletedCount;

        //DI API 연동용 내부 클래스
        public class ItemInformation
        {
            public int LineNum; //삭제한 행의 행번호
        }

        List<ItemInformation> itemInfoList = new List<ItemInformation>();
        ItemInformation itemInfo = new ItemInformation();

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM030.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM030_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM030");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.DataBrowser.BrowseBy = "DocNum"; //화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
                oDocNum = oFormDocEntry;

                oForm.Freeze(true);
                if (!string.IsNullOrEmpty(oFormDocEntry))
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                }
                PS_MM030_CreateItems();
                PS_MM030_ComboBox_Setting();
                PS_MM030_FormClear();
                PS_MM030_Add_MatrixRow(0, true);

                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1287", false); // 복제
                oForm.EnableMenu("1286", true); // 닫기
                oForm.EnableMenu("1285", false); // 복원
                oForm.EnableMenu("1284", true); // 취소
                oForm.EnableMenu("1293", true); // 행삭제

                g_deletedCount = 0;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                if (!string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_MM030_FormItemEnabled();
                    oForm.Items.Item("DocNum").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PS_MM030_Initialization();
                    PS_MM030_FormItemEnabled();
                }
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_MM030_CreateItems()
        {
            try
            {
                oDS_PS_MM030H = oForm.DataSources.DBDataSources.Item("@PS_MM030H");
                oDS_PS_MM030L = oForm.DataSources.DBDataSources.Item("@PS_MM030L");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

                oForm.DataSources.UserDataSources.Add("DocTotal", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("DocTotal").Specific.DataBind.SetBound(true, "", "DocTotal");

                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");

                oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");

                oDS_PS_MM030H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                oDS_PS_MM030H.SetValue("U_DueDate", 0, DateTime.Now.ToString("yyyyMMdd"));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM030_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                // 사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //품의형태
                sQry = "SELECT Code, Name From [@PSH_RETYPE]";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("POType").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //품의상태
                oForm.Items.Item("POStatus").Specific.ValidValues.Add("Y", "승인");
                oForm.Items.Item("POStatus").Specific.ValidValues.Add("N", "미승인");
                oDS_PS_MM030H.SetValue("U_POStatus", 0, "N");

                //품의종결 추가 20170807
                oForm.Items.Item("POFinish").Specific.ValidValues.Add("Y", "종결");
                oForm.Items.Item("POFinish").Specific.ValidValues.Add("N", "미종결");
                oDS_PS_MM030H.SetValue("U_POFinish", 0, "N");

                oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                //구매방식
                sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("Purchase").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oDS_PS_MM030H.SetValue("U_PurType", 0, "0");

                //지불조건
                sQry = "Select U_Minor, U_CdName From [@PS_SY001L] Where Code = 'M006' Order by U_Minor";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("Payment").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oDS_PS_MM030H.SetValue("U_Payment", 0, "20"); //기본값을 현금으로 변경(2014.07.23 송명규, 류석균 요청)

                //외주사유코드
                sQry = "  SELECT      T1.U_Minor AS [Code],";
                sQry += "             T1.U_CdName AS [Value]";
                sQry += " FROM        [@PS_SY001H] AS T0";
                sQry += "             INNER JOIN";
                sQry += "             [@PS_SY001L] AS T1";
                sQry += "                 ON T0.Code = T1.Code";
                sQry += " WHERE       T0.Code = 'P201'";
                sQry += "             AND T1.U_UseYN = 'Y'";
                sQry += " ORDER BY    T1.U_Seq";

                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("OutCode").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //통화
                sQry = "  SELECT      T0.U_Minor,";
                sQry += "             T1.CurrName + '(' + T0.U_CdName + ')'";
                sQry += " FROM        [@PS_SY001L] AS T0";
                sQry += "             LEFT JOIN";
                sQry += "             [OCRN] AS T1";
                sQry += "                 ON T0.U_Minor = T1.CurrCode";
                sQry += " WHERE       T0.Code = 'F004'";
                sQry += " ORDER BY    T0.U_Seq";

                dataHelpClass.Set_ComboList(oForm.Items.Item("DocCur").Specific, sQry, "", false, false);
                oForm.Items.Item("DocCur").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_MM030_Initialization
        /// </summary>
        private void PS_MM030_Initialization()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                
                oDS_PS_MM030H.SetValue("U_BPLId", 0, dataHelpClass.User_BPLID());//아이디별 사업장 세팅
                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD(); //아이디별 사번 세팅
                oForm.Items.Item("DocCur").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); //아이디별 부서 세팅
                oForm.Items.Item("DocRate").Specific.Value = 1;
                oForm.Items.Item("CardCode").Click();

                g_deletedCount = 0;
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
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM030_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("POType").Enabled = true;
                    oForm.Items.Item("Purchase").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oForm.Items.Item("POStatus").Enabled = false;
                    oForm.Items.Item("AdmsDate").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oMat01.Columns.Item("PQDocNum").Editable = true;
                    oMat01.Columns.Item("PQLinNum").Editable = true;
                    oMat01.Columns.Item("ItemCode").Editable = true;
                    oMat01.Columns.Item("ItemName").Editable = true;
                    oMat01.Columns.Item("OutSize").Editable = true;
                    oMat01.Columns.Item("OutUnit").Editable = true;
                    oMat01.Columns.Item("Qty").Editable = true;
                    oMat01.Columns.Item("UnWeight").Editable = true;
                    oMat01.Columns.Item("Price").Editable = true;
                    oMat01.Columns.Item("LinTotal").Editable = true;
                    oMat01.Columns.Item("WhsCode").Editable = true;
                    oMat01.Columns.Item("Comments").Editable = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("POType").Enabled = true;
                    oForm.Items.Item("Purchase").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oForm.Items.Item("POStatus").Enabled = false;
                    oForm.Items.Item("AdmsDate").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("POType").Enabled = false;
                    oForm.Items.Item("Purchase").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oForm.Items.Item("POStatus").Enabled = false;
                    oForm.Items.Item("AdmsDate").Enabled = false;
                    if (oForm.Items.Item("POFinish").Specific.Selected.Value == "Y")
                    {
                        oForm.Items.Item("POFinish").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("POFinish").Enabled = true;
                    }

                    if (oDS_PS_MM030H.GetValue("Canceled", 0).ToString().Trim() == "Y")
                    {
                        oForm.Items.Item("Mat01").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("Mat01").Enabled = true;
                        if (oDS_PS_MM030H.GetValue("U_POStatus", 0).ToString().Trim() == "Y" || oForm.Items.Item("POFinish").Specific.Selected.Value == "Y")
                        {
                            oMat01.Columns.Item("PQDocNum").Editable = false;
                            oMat01.Columns.Item("PQLinNum").Editable = false;
                            oMat01.Columns.Item("ItemCode").Editable = false;
                            oMat01.Columns.Item("ItemName").Editable = false;
                            oMat01.Columns.Item("OutSize").Editable = false;
                            oMat01.Columns.Item("OutUnit").Editable = false;
                            oMat01.Columns.Item("Qty").Editable = false;
                            oMat01.Columns.Item("UnWeight").Editable = false;
                            oMat01.Columns.Item("Price").Editable = false;
                            oMat01.Columns.Item("LinTotal").Editable = false;
                            oMat01.Columns.Item("WhsCode").Editable = false;
                            oMat01.Columns.Item("Comments").Editable = false;
                        }
                        else
                        {
                            oMat01.Columns.Item("PQDocNum").Editable = true;
                            oMat01.Columns.Item("PQLinNum").Editable = true;
                            oMat01.Columns.Item("ItemCode").Editable = true;
                            oMat01.Columns.Item("ItemName").Editable = true;
                            oMat01.Columns.Item("OutSize").Editable = true;
                            oMat01.Columns.Item("OutUnit").Editable = true;
                            oMat01.Columns.Item("Qty").Editable = true;
                            oMat01.Columns.Item("UnWeight").Editable = true;
                            oMat01.Columns.Item("Price").Editable = true;
                            oMat01.Columns.Item("LinTotal").Editable = true;
                            oMat01.Columns.Item("WhsCode").Editable = true;
                            oMat01.Columns.Item("Comments").Editable = true;
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
        /// PS_MM030_Add_MatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM030_Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM030L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM030L.Offset = oRow;
                oDS_PS_MM030L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_MM030_FormClear()
        {
            string DocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM030'", "");
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
            finally
            {
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_MM030_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int i;
            int sRow;
            double LineTotal;
            double TAmt;
            double FCPrice;
            double FCAmount;
            double DocRate;
            string sQry;
            string sSeq;
            string WhsCode;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet03 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sRow = oRow;
                switch (oUID)
                {
                    case "CardCode":
                        sQry = "Select CardName From OCRD Where CardCode = '" + oDS_PS_MM030H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);

                        oDS_PS_MM030H.SetValue("U_CardName", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());

                        sQry = "Select U_payment From OCRD Where CardCode = '" + oDS_PS_MM030H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);

                        oDS_PS_MM030H.SetValue("U_payment", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                        break;
                    case "CntcCode":
                        sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oDS_PS_MM030H.GetValue("U_CntcCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);

                        oDS_PS_MM030H.SetValue("U_CntcName", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                        break;
                    case "DocCur":
                    case "DocRate":
                        DocRate = Convert.ToDouble(oForm.Items.Item("DocRate").Specific.Value);

                        oMat01.FlushToDataSource();

                        for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                        {
                            if (oForm.Items.Item("DocCur").Specific.Value == "KRW")
                            {
                                FCPrice = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_Price", i).ToString().Trim());
                                FCAmount = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_LinTotal", i).ToString().Trim());
                            }
                            else
                            {
                                FCPrice = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_Price", i).ToString().Trim()) / DocRate;
                                FCAmount = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_Weight", i).ToString().Trim()) * FCPrice;
                            }

                            oDS_PS_MM030L.SetValue("U_FCPrice", i, Convert.ToString(FCPrice));  //외화단가
                            oDS_PS_MM030L.SetValue("U_FCAmount", i, Convert.ToString(FCAmount)); //외화금액
                        }
                        oMat01.LoadFromDataSource();
                        break;
                    case "Mat01":
                        if (oCol == "PQDocNum")
                        {
                            int chknum = 0;
                            if (int.TryParse(oMat01.Columns.Item("PQDocNum").Cells.Item(oRow).Specific.Value, out chknum) == false)
                            {
                                errMessage = "숫자만 입력하셔야합니다.";
                                throw new Exception();
                            }
                            oMat01.FlushToDataSource();

                            WhsCode = dataHelpClass.User_WhsCode("1");
                            WhsCode = Convert.ToString(Convert.ToDouble(codeHelpClass.Left(WhsCode, 2)) + oForm.Items.Item("BPLId").Specific.Value);

                            //프로시저로 수정(2012.03.26 송명규)
                            sQry = "      EXEC PS_MM030_50 '";
                            sQry += WhsCode.ToString().Trim() + "','";
                            sQry += oDS_PS_MM030L.GetValue("U_PQDocNum", oRow - 1).ToString().Trim() + "','";
                            sQry += oForm.Items.Item("DocCur").Specific.Selected.Value.ToString().Trim() + "'";

                            oRecordSet01.DoQuery(sQry);

                            if (oRecordSet01.RecordCount == 0)
                            {
                                errMessage = "구매견적문서가 취소되었거나 없습니다. 확인하세요.";
                                throw new Exception();
                            }
                            else
                            {
                                if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("PQDocNum").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_MM030_Add_MatrixRow(oMat01.RowCount, false);
                                    oMat01.Columns.Item("PQDocNum").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }

                                while (!oRecordSet01.EoF)
                                {
                                    sSeq = "Y";
                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                    {
                                        if (oDS_PS_MM030L.GetValue("U_PQDocNum", i).ToString().Trim() == oRecordSet01.Fields.Item(0).Value.ToString().Trim() && oDS_PS_MM030L.GetValue("U_PQLinNum", i).ToString().Trim() == oRecordSet01.Fields.Item(1).Value.ToString().Trim())
                                        {
                                            sSeq = "N";
                                        }
                                    }
                                    if (sSeq == "Y")
                                    {
                                        oDS_PS_MM030L.SetValue("U_PQDocNum", sRow - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_PQLinNum", sRow - 1, oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_ItemCode", sRow - 1, oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_ItemName", sRow - 1, oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_Qty", sRow - 1, oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_Weight", sRow - 1, oRecordSet01.Fields.Item("U_Weight").Value.ToString().Trim());
                                        if (Convert.ToDouble(oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim()) == 0)
                                        {
                                            oDS_PS_MM030L.SetValue("U_UnWeight", sRow - 1, "0");
                                        }
                                        else
                                        {
                                            oDS_PS_MM030L.SetValue("U_UnWeight", sRow - 1, Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item("U_Weight").Value.ToString().Trim()) / Convert.ToDouble(oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim())));
                                        }
                                        oDS_PS_MM030L.SetValue("U_WhsCode", sRow - 1, WhsCode);
                                        oDS_PS_MM030L.SetValue("U_WhsName", sRow - 1, oRecordSet01.Fields.Item("WhsName").Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_OutSize", sRow - 1, oRecordSet01.Fields.Item("U_OutSize").Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_OutUnit", sRow - 1, oRecordSet01.Fields.Item("U_OutUnit").Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_Auto", sRow - 1, oRecordSet01.Fields.Item("U_Auto").Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_Comments", sRow - 1, oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_ProcCode", sRow - 1, oRecordSet01.Fields.Item("U_ProcCode").Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_ProcName", sRow - 1, oRecordSet01.Fields.Item("U_ProcName").Value.ToString().Trim());
                                        //이론중량 추가(2012.03.26 송명규)
                                        oDS_PS_MM030L.SetValue("U_TWeight", sRow - 1, oRecordSet01.Fields.Item("U_TWeight").Value.ToString().Trim());

                                        if (oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "1" && oForm.Items.Item("Purchase").Specific.Value.ToString().Trim() == "30")
                                        {
                                            //창원 외주 단가
                                            sQry = "  Select      T0.U_eCardCod,";
                                            sQry += "             T0.U_ItemCode,";
                                            sQry += "             IsNull(T0.U_Cprice, 0) As U_Cprice,";
                                            sQry += "             T0.U_CtrDate ";
                                            sQry += " From        [@PS_PP006H] T0";
                                            sQry += "             Inner Join";
                                            sQry += "             (";
                                            sQry += "                 Select      U_eCardCod,";
                                            sQry += "                             U_ItemCode,";
                                            sQry += "                             U_CpCode,";
                                            sQry += "                             MAX(U_CtrDate) As U_CtrDate";
                                            sQry += "                 From        [@PS_PP006H]";
                                            sQry += "                 Group by    U_eCardCod,";
                                            sQry += "                             U_ItemCode,";
                                            sQry += "                             U_CpCode";
                                            sQry += "             ) T1 ";
                                            sQry += "                 On T1.U_eCardCod = T0.U_eCardCod";
                                            sQry += "                 And T1.U_ItemCode = T0.U_ItemCode";
                                            sQry += "                 And T1.U_CpCode = T0.U_CpCode";
                                            sQry += "                 And T1.U_CtrDate = T0.U_CtrDate";
                                            sQry += " Where       T0.U_eCardCod = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "' ";
                                            sQry += "             And T0.U_ItemCode = '" + oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim() + "' ";
                                            sQry += "             And T0.U_CpCode = '" + oRecordSet01.Fields.Item("U_ProcCode").Value.ToString().Trim() + "' ";

                                            oRecordSet02.DoQuery(sQry);

                                            oDS_PS_MM030L.SetValue("U_Price", sRow - 1, oRecordSet02.Fields.Item("U_Cprice").Value.ToString().Trim());
                                            oDS_PS_MM030L.SetValue("U_LinTotal", sRow - 1, Convert.ToDouble(Convert.ToDouble(oRecordSet02.Fields.Item("U_Cprice").Value)) * Convert.ToDouble(oRecordSet01.Fields.Item("U_Weight").Value.ToString().Trim()));
                                        }
                                        //부산 외주단가 추가(2023.07.28 박우근 수정)
                                        else if (oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "2" && oForm.Items.Item("Purchase").Specific.Value.ToString().Trim() == "30")
                                        {
                                            //부산 외주 단가
                                            sQry = "  Select      T0.U_eCardCod,";
                                            sQry += "             T0.U_ItemCode,";
                                            sQry += "             IsNull(T0.U_Cprice, 0) As U_Cprice,";
                                            sQry += "             T0.U_CtrDate ";
                                            sQry += " From        [@PS_PP006H] T0";
                                            sQry += "             Inner Join";
                                            sQry += "             (";
                                            sQry += "                 Select      U_eCardCod,";
                                            sQry += "                             U_ItemCode,";
                                            sQry += "                             U_CpCode,";
                                            sQry += "                             MAX(U_CtrDate) As U_CtrDate";
                                            sQry += "                 From        [@PS_PP006H]";
                                            sQry += "                 Group by    U_eCardCod,";
                                            sQry += "                             U_ItemCode,";
                                            sQry += "                             U_CpCode";
                                            sQry += "             ) T1 ";
                                            sQry += "                 On T1.U_eCardCod = T0.U_eCardCod";
                                            sQry += "                 And T1.U_ItemCode = T0.U_ItemCode";
                                            sQry += "                 And T1.U_CpCode = T0.U_CpCode";
                                            sQry += "                 And T1.U_CtrDate = T0.U_CtrDate";
                                            sQry += " Where       T0.U_eCardCod = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "' ";
                                            sQry += "             And T0.U_ItemCode = '" + oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim() + "' ";
                                            sQry += "             And T0.U_CpCode = '" + oRecordSet01.Fields.Item("U_ProcCode").Value.ToString().Trim() + "' ";

                                            oRecordSet02.DoQuery(sQry);

                                            oDS_PS_MM030L.SetValue("U_Price", sRow - 1, oRecordSet02.Fields.Item("U_Cprice").Value.ToString().Trim());
                                            oDS_PS_MM030L.SetValue("U_LinTotal", sRow - 1, Convert.ToDouble(Convert.ToDouble(oRecordSet02.Fields.Item("U_Cprice").Value)) * Convert.ToDouble(oRecordSet01.Fields.Item("U_Weight").Value.ToString().Trim()));
                                        }

                                        //구매실적단가 가져오기
                                        sQry = "EXEC PS_MM030_03 '";
                                        sQry += oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "','";
                                        sQry += oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim() + "','";
                                        sQry += oRecordSet01.Fields.Item("U_ProcCode").Value.ToString().Trim() + "','";
                                        sQry += oForm.Items.Item("DocNum").Specific.Value.ToString().Trim() + "','";
                                        sQry += oForm.Items.Item("DocDate").Specific.Value.ToString().Trim() + "'";

                                        oRecordSet03.DoQuery(sQry);

                                        oDS_PS_MM030L.SetValue("U_LCardCd", sRow - 1, oRecordSet03.Fields.Item("CardCode").Value.ToString().Trim());
                                        oDS_PS_MM030L.SetValue("U_LCardNm", sRow - 1, oRecordSet03.Fields.Item("U_CardName").Value.ToString().Trim());

                                        if (Convert.ToInt32(oRecordSet03.Fields.Item("U_DocDate").Value.ToString("yyyyMMdd")) < 19000102)
                                        {
                                            oDS_PS_MM030L.SetValue("U_LDocDate", sRow - 1, "");
                                        }
                                        else
                                        {
                                            oDS_PS_MM030L.SetValue("U_LDocDate", sRow - 1, oRecordSet03.Fields.Item("U_DocDate").Value.ToString("yyyyMMdd"));
                                        }
                                        oDS_PS_MM030L.SetValue("U_LPrice", sRow - 1, oRecordSet03.Fields.Item("U_Price").Value.ToString().Trim());
                                        //구매실적단가 가져오기 끝

                                        //외주사유관련 내용 시작(2013.08.13 송명규 수정)
                                        if (!string.IsNullOrEmpty(oRecordSet01.Fields.Item("U_OutCode").Value.ToString().Trim()))
                                        {
                                            oDS_PS_MM030L.SetValue("U_OutCode", sRow - 1, oRecordSet01.Fields.Item("U_OutCode").Value.ToString().Trim()); //외주사유코드
                                        }
                                        oDS_PS_MM030L.SetValue("U_OutNote", sRow - 1, oRecordSet01.Fields.Item("U_OutNote").Value.ToString().Trim()); //외주사유내용
                                        oDS_PS_MM030L.SetValue("U_MComment", sRow - 1, oRecordSet01.Fields.Item("U_MComment").Value.ToString().Trim()); //자재담당의견
                                        //외주사유관련 내용 종료(2013.08.13 송명규 수정)

                                        //환율 세팅 시작(2016.06.22 송명규 수정)
                                        oForm.Items.Item("DocRate").Specific.Value = oRecordSet01.Fields.Item("DocRate").Value.ToString().Trim(); //환율
                                        //환율 세팅 종료(2016.06.22 송명규 수정)

                                        PS_MM030_Add_MatrixRow(sRow, false);
                                        sRow += 1;
                                    }
                                    oRecordSet01.MoveNext();
                                }

                                if (oMat01.VisualRowCount > 0)
                                {
                                    if (string.IsNullOrEmpty(oDS_PS_MM030L.GetValue("U_ItemCode", oMat01.VisualRowCount - 1).ToString().Trim()))
                                    {
                                        oDS_PS_MM030L.RemoveRecord(oMat01.VisualRowCount - 1);
                                    }
                                }

                                oMat01.LoadFromDataSource();

                                PS_MM030_TotalAmount_Calculate();
                            }
                        }
                        else if (oCol == "ItemCode")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_MM030_Add_MatrixRow(oMat01.RowCount, false);
                                }
                            }
                            sQry = "Select ItemName, FrgnName From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (oCol == "Qty")
                        {
                            oMat01.FlushToDataSource();
                            if (Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(oRow).Specific.Value.ToString().Trim()) == 0)
                            {
                                oDS_PS_MM030L.SetValue("U_LinTotal", oRow - 1, "0");
                            }
                            else
                            {
                                oDS_PS_MM030L.SetValue("U_LinTotal", oRow - 1, Convert.ToString(System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(oRow).Specific.Value.ToString().Trim()) * Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(oRow).Specific.Value.ToString().Trim()), 0)));
                            }
                            oMat01.LoadFromDataSource();

                            PS_MM030_TotalAmount_Calculate();

                            oMat01.Columns.Item("UnWeight").Cells.Item(oRow).Click();
                        }
                        else if (oCol == "Price")
                        {
                            oMat01.FlushToDataSource();
                            if (Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(oRow).Specific.Value.ToString().Trim()) == 0)
                            {
                                oDS_PS_MM030L.SetValue("U_LinTotal", oRow - 1, "0");
                            }
                            else
                            {
                                //품목분류에 따라 이론중량에 따른 이론금액을 계산(2012.03.26 송명규 추가)
                                sQry = "SELECT U_ItmMsort FROM OITM WHERE ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);

                                if (oRecordSet01.Fields.Item("U_ItmMsort").Value == "30603") //품목의 분류가 원재_소재류 중 봉인 경우
                                {
                                    LineTotal = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("TWeight").Cells.Item(oRow).Specific.Value.ToString().Trim()) * Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(oRow).Specific.Value.ToString().Trim()), 0);
                                    TAmt = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("TWeight").Cells.Item(oRow).Specific.Value.ToString().Trim()) * Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(oRow).Specific.Value.ToString().Trim()), 0);
                                }
                                else //품목의 분류가 원재_소재류 중 봉을 제외한 경우
                                {
                                    LineTotal = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(oRow).Specific.Value.ToString().Trim()) * Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(oRow).Specific.Value.ToString().Trim()), 0);
                                    TAmt = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(oRow).Specific.Value.ToString().Trim()) * Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(oRow).Specific.Value.ToString().Trim()), 0);
                                }

                                //외화단가 계산
                                DocRate = Convert.ToDouble(oForm.Items.Item("DocRate").Specific.Value);
                                if (oForm.Items.Item("DocCur").Specific.Value == "KRW")
                                {
                                    FCPrice = Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(oRow).Specific.Value.ToString().Trim());
                                    FCAmount = Convert.ToDouble(oMat01.Columns.Item("LineTotal").Cells.Item(oRow).Specific.Value.ToString().Trim());
                                }
                                else
                                {
                                    FCPrice = Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(oRow).Specific.Value.ToString().Trim()) / DocRate;
                                    FCAmount = Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(oRow).Specific.Value.ToString().Trim()) * FCPrice;
                                }
                                oDS_PS_MM030L.SetValue("U_LinTotal", oRow - 1, Convert.ToString(LineTotal)); //금액
                                oDS_PS_MM030L.SetValue("U_TAmt", oRow - 1, Convert.ToString(LineTotal)); //이론금액
                                oDS_PS_MM030L.SetValue("U_FCPrice", oRow - 1, Convert.ToString(FCPrice)); //외화단가
                                oDS_PS_MM030L.SetValue("U_FCAmount", oRow - 1, Convert.ToString(FCAmount)); //외화금액
                            }
                            oMat01.LoadFromDataSource();

                            if (LastPick == "")
                            {
                                PS_MM030_TotalAmount_Calculate();
                            }
                            oMat01.Columns.Item("Price").Cells.Item(oRow).Click();
                        }
                        else if (oCol == "LinTotal")
                        {
                            PS_MM030_TotalAmount_Calculate();

                            oMat01.Columns.Item("Price").Cells.Item(oRow).Click();
                        }
                        else if (oCol == "WhsCode")
                        {
                            sQry = "Select WhsName From [OWHS] Where WhsCode = '" + oMat01.Columns.Item("WhsCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);

                            oMat01.Columns.Item("WhsName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        oMat01.AutoResizeColumns();
                        break;
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet03);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_MM030_HeaderSpaceLineDel()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_DocDate", 0)))
                {
                    errMessage = "전기일은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                // 마감일자 Check
                //else if (dataHelpClass.Check_Finish_Status(oDS_PS_MM030H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_MM030H.GetValue("U_DocDate", 0).ToString().Trim().Substring(0, 6)) == false)
                //{
                //    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.";
                //    throw new Exception();
                //}
                else if (string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_CardCode", 0).ToString().Trim()))
                {
                    errMessage = "거래처는 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_CntcCode", 0)))
                {
                    errMessage = "담당자는 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_POType", 0)))
                {
                    errMessage = "품의형태는 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_Purchase", 0)))
                {
                    errMessage = "구매방식은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_DueDate", 0)))
                {
                    errMessage = "납품일은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (oDS_PS_MM030H.GetValue("Canceled", 0) == "Y")
                {
                    errMessage = "문서 상태가 취소입니다. 수정할 수 없습니다.";
                    throw new Exception();
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM030_MatrixSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM030_MatrixSpaceLineDel()
        {
            bool returnValue = false;
            int i;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM030L.GetValue("U_ItemCode", i)))
                    {
                        errMessage = "품목코드는 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_MM030L.GetValue("U_Weight", i)))
                    {
                        errMessage = "중량은 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (Convert.ToDouble(oDS_PS_MM030L.GetValue("U_Price", i)) == 0)
                    {
                        errMessage = "단가는 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (Convert.ToDouble(oDS_PS_MM030L.GetValue("U_LinTotal", i)) == 0)
                    {
                        errMessage = "금액은 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_MM030L.GetValue("U_WhsCode", i)))
                    {
                        errMessage = "창고코드는 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (PS_MM030_CheckDate(oMat01.Columns.Item("PQDocNum").Cells.Item(i + 1).Specific.Value) == false) //구매견적과 일자 체크
                    {
                        errMessage = i + 1 + "행 [" + oMat01.Columns.Item("ItemCode").Cells.Item(i + 1).Specific.Value + "]의 구매품의일은 구매견적일과 같거나 늦어야합니다. 확인하십시오. 해당 품의는 전체가 등록되지 않습니다";
                        throw new Exception();
                    }
                }
                oMat01.LoadFromDataSource();
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM030_Delete_EmptyRow()
        /// </summary>
        private void PS_MM030_Delete_EmptyRow()
        {
            int i;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();

                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM030L.GetValue("U_ItemCode", i).ToString().Trim()))
                    {
                        oDS_PS_MM030L.RemoveRecord(i); // Mat01에 마지막라인(빈라인) 삭제
                    }
                }
                oMat01.LoadFromDataSource();
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
        }

        /// <summary>
        /// PS_MM030_oPurchaseOrders_Add
        /// </summary>
        /// <returns></returns>
        private bool PS_MM030_oPurchaseOrders_Add()
        {
            bool returnValue = false;
            int i;
            int RetVal;
            int errDICode;
            string errDIMsg;
            string DocEntry;
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Documents DI_oPurchaseOrders = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_PODocNum", 0).ToString().Trim()))
                {
                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                    PSH_Globals.oCompany.StartTransaction();
                    oMat01.FlushToDataSource();

                    DI_oPurchaseOrders.CardCode = oForm.Items.Item("CardCode").Specific.Value;
                    DI_oPurchaseOrders.BPL_IDAssignedToInvoice = Convert.ToInt32(oDS_PS_MM030H.GetValue("U_BPLId", 0).ToString().Trim());
                    DI_oPurchaseOrders.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                    DI_oPurchaseOrders.DocDueDate = DateTime.ParseExact(oForm.Items.Item("DueDate").Specific.Value, "yyyyMMdd", null);
                    DI_oPurchaseOrders.DocCurrency = oForm.Items.Item("DocCur").Specific.Selected.Value; //ISO 통화기호 추가(2020.04.06 송명규)
                    DI_oPurchaseOrders.DocRate = Convert.ToDouble(oForm.Items.Item("DocRate").Specific.Value); //환율 추가(2020.04.06 송명규)
                    DI_oPurchaseOrders.Comments = oForm.Items.Item("Comments").Specific.Value;
                    DI_oPurchaseOrders.UserFields.Fields.Item("U_reType").Value = oForm.Items.Item("POType").Specific.Selected.Value;
                    DI_oPurchaseOrders.UserFields.Fields.Item("U_okYN").Value = oForm.Items.Item("POStatus").Specific.Selected.Value;
                    DI_oPurchaseOrders.UserFields.Fields.Item("U_OrdTyp").Value = oForm.Items.Item("Purchase").Specific.Selected.Value;

                    sQry = "Select ECVatGroup From [OCRD] Where CardCode = '" + oDS_PS_MM030H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (oDS_PS_MM030H.GetValue("U_Purchase", 0).ToString().Trim() == "30" || oDS_PS_MM030H.GetValue("U_Purchase", 0).ToString().Trim() == "40" || oDS_PS_MM030H.GetValue("U_Purchase", 0).ToString().Trim() == "60")
                    {
                        DI_oPurchaseOrders.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                        for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                        {
                            if (i > 0)
                            {
                                DI_oPurchaseOrders.Lines.Add();
                            }
                            DI_oPurchaseOrders.Lines.SetCurrentLine(i);

                            DI_oPurchaseOrders.Lines.ItemDescription = oDS_PS_MM030L.GetValue("U_ItemName", i).ToString().Trim() + "-" + oDS_PS_MM030L.GetValue("U_OutSize", i).ToString().Trim() + "-" + oDS_PS_MM030L.GetValue("U_OutUnit", i).ToString().Trim();
                            //외화처리 기능 구헌(2020.04.06 송명규)_S
                            if (oDS_PS_MM030H.GetValue("U_DocCur", 0).ToString().Trim() == "KRW")
                            {
                                DI_oPurchaseOrders.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_LinTotal", i).ToString().Trim());
                            }
                            else
                            {
                                DI_oPurchaseOrders.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_FCAmount", i).ToString().Trim());
                            }
                            //외화처리 기능 구헌(2020.04.06 송명규)_E
                            DI_oPurchaseOrders.Lines.VatGroup = oRecordSet01.Fields.Item("ECVatGroup").Value;
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sItemCode").Value = oDS_PS_MM030L.GetValue("U_ItemCode", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sItemName").Value = oDS_PS_MM030L.GetValue("U_ItemName", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sSize").Value = oDS_PS_MM030L.GetValue("U_OutSize", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sUnit").Value = oDS_PS_MM030L.GetValue("U_OutUnit", i).ToString().Trim();
                            if (!string.IsNullOrEmpty(oDS_PS_MM030L.GetValue("U_Qty", i).ToString().Trim()) || Convert.ToDouble(oDS_PS_MM030L.GetValue("U_Qty", i).ToString().Trim()) != 0)
                            {
                                DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sQty").Value = oDS_PS_MM030L.GetValue("U_Qty", i).ToString().Trim();
                            }
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sWeight").Value = oDS_PS_MM030L.GetValue("U_Weight", i).ToString().Trim();

                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_MM010Doc").Value = oDS_PS_MM030L.GetValue("U_PQDocNum", i).ToString().Trim(); //구매견적문서번호(2017.04.13 송명규 추가)
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_MM010Lin").Value = oDS_PS_MM030L.GetValue("U_PQLinNum", i).ToString().Trim(); //구매견적라인번호(2017.04.13 송명규 추가)

                            //작번을 입력
                            //If Trim(oDS_PS_MM030H.GetValue("U_Purchase", 0)) = "10" Then
                            sQry = "EXEC PS_MM030_04 '" + oDS_PS_MM030L.GetValue("U_PQDocNum", i).ToString().Trim() + "', '" + oDS_PS_MM030L.GetValue("U_PQLinNum", i).ToString().Trim() + "'";
                            oRecordSet02.DoQuery(sQry);

                            //작번
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_OrdNum").Value = oRecordSet02.Fields.Item(0).Value.ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_OrdSub1").Value = oRecordSet02.Fields.Item(1).Value.ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_OrdSub2").Value = oRecordSet02.Fields.Item(2).Value.ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_Payment").Value = oDS_PS_MM030H.GetValue("U_Payment", 0).ToString().Trim();
                        }
                    }
                    else
                    {
                        DI_oPurchaseOrders.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                        for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                        {
                            if (i > 0)
                            {
                                DI_oPurchaseOrders.Lines.Add();
                            }
                            DI_oPurchaseOrders.Lines.SetCurrentLine(i);

                            DI_oPurchaseOrders.Lines.ItemCode = oDS_PS_MM030L.GetValue("U_ItemCode", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.Quantity = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_Weight", i).ToString().Trim());

                            //외화처리 기능 구헌(2020.04.06 송명규)_S
                            if (oDS_PS_MM030H.GetValue("U_DocCur", 0).ToString().Trim() == "KRW")
                            {
                                DI_oPurchaseOrders.Lines.Price = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_Price", i).ToString().Trim());
                                DI_oPurchaseOrders.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_LinTotal", i).ToString().Trim());
                            }
                            else
                            {
                                DI_oPurchaseOrders.Lines.Price = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_FCPrice", i).ToString().Trim());
                                DI_oPurchaseOrders.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_FCAmount", i).ToString().Trim());
                            }
                            //외화처리 기능 구헌(2020.04.06 송명규)_E
                            DI_oPurchaseOrders.Lines.WarehouseCode = oDS_PS_MM030L.GetValue("U_WhsCode", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.VatGroup = oRecordSet01.Fields.Item("ECVatGroup").Value;
                            if (!string.IsNullOrEmpty(oDS_PS_MM030L.GetValue("U_Qty", i)) || Convert.ToDouble(oDS_PS_MM030L.GetValue("U_Qty", i)) != 0)
                            {
                                DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_Qty").Value = oDS_PS_MM030L.GetValue("U_Qty", i).ToString().Trim();
                            }
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_Weight").Value = oDS_PS_MM030L.GetValue("U_Weight", i).ToString().Trim();

                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_MM010Doc").Value = oDS_PS_MM030L.GetValue("U_PQDocNum", i).ToString().Trim();
                            //구매견적문서번호(2017.04.13 송명규 추가)
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_MM010Lin").Value = oDS_PS_MM030L.GetValue("U_PQLinNum", i).ToString().Trim();
                            //구매견적라인번호(2017.04.13 송명규 추가)

                            sQry = "EXEC PS_MM030_04 '" + oDS_PS_MM030L.GetValue("U_PQDocNum", i).ToString().Trim() + "', '" + oDS_PS_MM030L.GetValue("U_PQLinNum", i).ToString().Trim() + "'";
                            oRecordSet02.DoQuery(sQry);

                            //작번
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_OrdNum").Value = oRecordSet02.Fields.Item(0).Value.ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_OrdSub1").Value = oRecordSet02.Fields.Item(1).Value.ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_OrdSub2").Value = oRecordSet02.Fields.Item(2).Value.ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_Payment").Value = oDS_PS_MM030H.GetValue("U_Payment", 0).ToString().Trim();
                        }
                    }

                    RetVal = DI_oPurchaseOrders.Add(); //완료
                    if (0 != RetVal)
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
                    }
                    else
                    {
                        PSH_Globals.oCompany.GetNewObjectCode(out DocEntry);
                        oDS_PS_MM030H.SetValue("U_PODocNum", 0, DocEntry);

                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                    returnValue = true;
                }
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oPurchaseOrders);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM030_oPurchaseOrders_Update
        /// </summary>
        /// <returns></returns>
        private bool PS_MM030_oPurchaseOrders_Update()
        {
            bool returnValue = false;
            int i;
            int errDICode;
            int RetVal;
            string sQry;
            string errDIMsg;
            string DocDate;
            string DueDate;
            string DocEntry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Documents DI_oPurchaseOrders = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

            try
            {
                oMat01.FlushToDataSource();

                DocEntry = oDS_PS_MM030H.GetValue("U_PODocNum", 0).ToString().Trim();
                DocDate = oDS_PS_MM030H.GetValue("U_DocDate", 0);
                DueDate = oDS_PS_MM030H.GetValue("U_DueDate", 0);

                if (!string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_PODocNum", 0).ToString().Trim()))
                {
                    if (DI_oPurchaseOrders.GetByKey(Convert.ToInt32(DocEntry)) == false)
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
                    }

                    DI_oPurchaseOrders.CardCode = oForm.Items.Item("CardCode").Specific.Value;
                    DI_oPurchaseOrders.BPL_IDAssignedToInvoice = Convert.ToInt32(oDS_PS_MM030H.GetValue("U_BPLId", 0).ToString().Trim());
                    DI_oPurchaseOrders.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                    DI_oPurchaseOrders.DocDueDate = DateTime.ParseExact(oForm.Items.Item("DueDate").Specific.Value, "yyyyMMdd", null);
                    DI_oPurchaseOrders.DocCurrency = oForm.Items.Item("DocCur").Specific.Selected.Value; //ISO 통화기호 추가(2020.04.06 송명규)
                    DI_oPurchaseOrders.DocRate = Convert.ToDouble(oForm.Items.Item("DocRate").Specific.Value); //환율 추가(2020.04.06 송명규)
                    DI_oPurchaseOrders.Comments = oForm.Items.Item("Comments").Specific.Value;
                    DI_oPurchaseOrders.UserFields.Fields.Item("U_reType").Value = oForm.Items.Item("POType").Specific.Selected.Value;
                    DI_oPurchaseOrders.UserFields.Fields.Item("U_okYN").Value = oForm.Items.Item("POStatus").Specific.Selected.Value;
                    DI_oPurchaseOrders.UserFields.Fields.Item("U_OrdTyp").Value = oForm.Items.Item("Purchase").Specific.Selected.Value;

                    sQry = "Select ECVatGroup From [OCRD] Where CardCode = '" + oDS_PS_MM030H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (oDS_PS_MM030H.GetValue("U_Purchase", 0).ToString().Trim() == "30" || oDS_PS_MM030H.GetValue("U_Purchase", 0).ToString().Trim() == "40" || oDS_PS_MM030H.GetValue("U_Purchase", 0).ToString().Trim() == "60")
                    {
                        DI_oPurchaseOrders.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                        for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                        {
                            if (i > 0)
                            {
                                DI_oPurchaseOrders.Lines.Add();
                            }
                            DI_oPurchaseOrders.Lines.SetCurrentLine(i);
                            DI_oPurchaseOrders.Lines.ItemDescription = oDS_PS_MM030L.GetValue("U_ItemName", i).ToString().Trim() + "-" + oDS_PS_MM030L.GetValue("U_OutSize", i).ToString().Trim() + "-" + oDS_PS_MM030L.GetValue("U_OutUnit", i).ToString().Trim();
                            //외화처리 기능 구헌(2020.04.06 송명규)_S
                            if (oDS_PS_MM030H.GetValue("U_DocCur", 0).ToString().Trim() == "KRW")
                            {
                                DI_oPurchaseOrders.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_LinTotal", i).ToString().Trim());
                            }
                            else
                            {
                                DI_oPurchaseOrders.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_FCAmount", i).ToString().Trim());
                            }
                            //외화처리 기능 구헌(2020.04.06 송명규)_E
                            DI_oPurchaseOrders.Lines.ShipDate = DateTime.ParseExact(oForm.Items.Item("DueDate").Specific.Value, "yyyyMMdd", null);
                            DI_oPurchaseOrders.Lines.VatGroup = oRecordSet01.Fields.Item("ECVatGroup").Value;
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sItemCode").Value = oDS_PS_MM030L.GetValue("U_ItemCode", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sItemName").Value = oDS_PS_MM030L.GetValue("U_ItemName", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sSize").Value = oDS_PS_MM030L.GetValue("U_OutSize", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sUnit").Value = oDS_PS_MM030L.GetValue("U_OutUnit", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sQty").Value = oDS_PS_MM030L.GetValue("U_Qty", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_sWeight").Value = oDS_PS_MM030L.GetValue("U_Weight", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_Payment").Value = oDS_PS_MM030H.GetValue("U_Payment", 0).ToString().Trim();
                        }
                    }
                    else
                    {
                        DI_oPurchaseOrders.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                        for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                        {
                            if (i > 0)
                            {
                                DI_oPurchaseOrders.Lines.Add();
                            }
                            DI_oPurchaseOrders.Lines.SetCurrentLine(i);
                            DI_oPurchaseOrders.Lines.ItemCode = oDS_PS_MM030L.GetValue("U_ItemCode", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.Quantity = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_Weight", i).ToString().Trim());
                            //외화처리 기능 구헌(2020.04.06 송명규)_S
                            if (oDS_PS_MM030H.GetValue("U_DocCur", 0).ToString().Trim() == "KRW")
                            {
                                DI_oPurchaseOrders.Lines.Price = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_Price", i).ToString().Trim());
                                DI_oPurchaseOrders.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_LinTotal", i).ToString().Trim());
                            }
                            else
                            {
                                DI_oPurchaseOrders.Lines.Price = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_FCPrice", i).ToString().Trim());
                                DI_oPurchaseOrders.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_FCAmount", i).ToString().Trim());
                            }
                            DI_oPurchaseOrders.Lines.DiscountPercent = 0;
                            DI_oPurchaseOrders.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM030L.GetValue("U_LinTotal", i).ToString().Trim());
                            DI_oPurchaseOrders.Lines.WarehouseCode = oDS_PS_MM030L.GetValue("U_WhsCode", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.VatGroup = oRecordSet01.Fields.Item("ECVatGroup").Value.ToString().Trim();
                            DI_oPurchaseOrders.Lines.ShipDate = DateTime.ParseExact(oForm.Items.Item("DueDate").Specific.Value, "yyyyMMdd", null);
                            if (!string.IsNullOrEmpty(oDS_PS_MM030L.GetValue("U_Qty", i).ToString().Trim()) || Convert.ToDouble(oDS_PS_MM030L.GetValue("U_Qty", i).ToString().Trim()) != 0)
                            {
                                DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_Qty").Value = oDS_PS_MM030L.GetValue("U_Qty", i).ToString().Trim();
                            }
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_Weight").Value = oDS_PS_MM030L.GetValue("U_Weight", i).ToString().Trim();
                            DI_oPurchaseOrders.Lines.UserFields.Fields.Item("U_Payment").Value = oDS_PS_MM030H.GetValue("U_Payment", 0).ToString().Trim();
                        }
                    }
                    
                    RetVal = DI_oPurchaseOrders.Update(); //완료

                    if (0 != RetVal)
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oPurchaseOrders);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM030_oPurchaseOrders_Cancel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM030_oPurchaseOrders_Cancel()
        {
            bool returnValue = false;
            int RetVal;
            int errDICode;
            string errDIMsg;
            string DocEntry;
            string errMessage = string.Empty;
            SAPbobsCOM.Documents DI_oPurchaseOrders = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

            try
            {
                DocEntry = oDS_PS_MM030H.GetValue("U_PODocNum", 0).ToString().Trim();

                if (!string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_PODocNum", 0).ToString().Trim()))
                {
                    if (DI_oPurchaseOrders.GetByKey(Convert.ToInt32(DocEntry)) == false) //완료
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
                    }
                    RetVal = DI_oPurchaseOrders.Cancel();
                    if (0 != RetVal)
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oPurchaseOrders);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM030_oPurchaseOrders_Close
        /// </summary>
        /// <returns></returns>
        private bool PS_MM030_oPurchaseOrders_Close()
        {
            bool returnValue = false;
            int RetVal;
            int errDICode;
            string errDIMsg;
            string DocEntry;
            string errMessage = string.Empty;
            SAPbobsCOM.Documents DI_oPurchaseOrders = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

            try
            {
                DocEntry = oDS_PS_MM030H.GetValue("U_PODocNum", 0).ToString().Trim();

                if (!string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_PODocNum", 0).ToString().Trim()))
                {
                    if (DI_oPurchaseOrders.GetByKey(Convert.ToInt32(DocEntry)) == false) //완료
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
                    }
                    RetVal = DI_oPurchaseOrders.Close();
                    if (0 != RetVal)
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oPurchaseOrders);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM030_oPurchaseOrders_Close
        /// </summary>
        /// <returns></returns>
        private bool PS_MM030_oPurchaseOrders_LineDelete(int pLineNum)
        {
            bool returnValue = false;
            int RetVal;
            int errDICode;
            string errDIMsg;
            string DocEntry;
            string errMessage = string.Empty;
            SAPbobsCOM.Documents DI_oPurchaseOrders = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

            try
            {
                DocEntry = oDS_PS_MM030H.GetValue("U_PODocNum", 0).ToString().Trim();
                
                if (DI_oPurchaseOrders.GetByKey(Convert.ToInt32(DocEntry)) == false) //완료
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                    throw new Exception();
                }
                else
                {
                    DI_oPurchaseOrders.Lines.SetCurrentLine(pLineNum - 1); //LineID에서 1을 빼야 구매오더 행번호와 일치
                    DI_oPurchaseOrders.Lines.Delete();
                }

                RetVal = DI_oPurchaseOrders.Update();

                if (0 != RetVal)
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                    throw new Exception();
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oPurchaseOrders);
            }
            return returnValue;
        }

        /// <summary>
        /// Report_Export
        /// </summary>
        [STAThread]
        private void PS_MM030_Print_Report01()
        {
            string DocNum;
            string WinTitle;
            string ReportName;
            string BPLId;
            string errMessage = string.Empty;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocNum = oForm.Items.Item("DocNum").Specific.Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

                WinTitle = "주문서 [PS_MM030_01]";

                if (BPLId == "2")
                {
                    ReportName = "PS_MM030_03.rpt"; //기계사업부(2015.08.11 추가)
                }
                else
                {
                    ReportName = "PS_MM030_01.rpt";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                // Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocNum", DocNum));

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
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
        }

        /// <summary>
        /// Report_Export
        /// </summary>
        [STAThread]
        private void PS_MM030_Print_Report02()
        {
            string WinTitle;
            string ReportName;
            string TitleName = string.Empty;
            string DocNum;
            string Purchase;
            string errMessage = string.Empty;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocNum = oForm.Items.Item("DocNum").Specific.Value.ToString().Trim();
                Purchase = oForm.Items.Item("Purchase").Specific.Value.ToString().Trim();
                WinTitle = "품의서 [PS_MM030_02]";
                ReportName = "PS_MM030_02.rpt";

                //Formula 수식필드
                if (Purchase == "10")
                {
                    TitleName = "구 매 품 의 서 (원자재)";
                }
                else if (Purchase == "20")
                {
                    TitleName = "구 매 품 의 서 (부재료)";
                }
                else if (Purchase == "30")
                {
                    TitleName = "구 매 품 의 서 (외주가공비)";
                }
                else if (Purchase == "40")
                {
                    TitleName = "구 매 품 의 서 (외주제작)";
                }
                else if (Purchase == "50")
                {
                    TitleName = "구 매 품 의 서 (상품)";
                }
                else if (Purchase == "60")
                {
                    TitleName = "구 매 품 의 서 (고정자산)";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

                // Formula
                dataPackFormula.Add(new PSH_DataPackClass("@F01", TitleName));

                // Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocNum", DocNum));

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_MM030_Check_Purchase_Type(int prmPQDocNum, string prmPOType)
        {
            bool returnValue = false;
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT     U_Purchase";
                sQry += " FROM       [@PS_MM010H]";
                sQry += " WHERE      DocEntry = " + prmPQDocNum;

                oRecordSet01.DoQuery(sQry);

                if (oRecordSet01.Fields.Item("U_Purchase").Value == prmPOType)
                {
                    returnValue = true;
                }
                else
                {
                    returnValue = false;
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM030_CheckDate
        /// </summary>
        /// <returns></returns>
        private bool PS_MM030_CheckDate(string pBaseEntry)
        {
            bool returnValue = false;
            string Query01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "EXEC PS_Z_CHECK_DATE '";
                Query01 += pBaseEntry + "','"; // BaseEntry
                Query01 += "" + "','";  //BaseLine
                Query01 += "PS_MM030" + "','"; //DocType
                Query01 += oForm.Items.Item("DocDate").Specific.Value.ToString().Trim() + "'"; //CurDocDate

                oRecordSet01.DoQuery(Query01);

                if (oRecordSet01.Fields.Item("ReturnValue").Value == "False")
                {
                    returnValue = false;
                }
                else
                {
                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return returnValue;
        }

        /// <summary>
        /// 최종구매단가 -> 단가(원) 복사
        /// </summary>
        /// <returns></returns>
        private void PS_MM030_Copy_Price()
        {
            int i;
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                ProgressBar01.Text = "복사 시작!";
                oForm.Freeze(true);

                LastPick = "Copy";
                for (i = 1; i < oMat01.RowCount; i++)
                {
                    oMat01.Columns.Item("Price").Cells.Item(i).Specific.Value = oMat01.Columns.Item("LPrice").Cells.Item(i).Specific.Value;

                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + Convert.ToString(oMat01.VisualRowCount - 1) + "건 처리중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) // 데이터 수정시 후 갱신모드로 변경이 원활하지 않아 강제로 업데이트모드로 변경
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                ProgressBar01.Text = "수량, 중량, 금액합계 계산중!";
                PS_MM030_TotalAmount_Calculate();
                LastPick = "";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_MM030_LineDelete_Possible
        /// </summary>
        /// <param name="pDocEntry"></param>
        /// <param name="pLineID"></param>
        /// <returns></returns>
        private bool PS_MM030_LineDelete_Possible(string pDocEntry, string pLineID)
        {
            bool returnValue = false;
            string QueryStr01;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                QueryStr01 = "  SELECT      COUNT(*) AS [Cnt]";
                QueryStr01 += " FROM        [@PS_MM050H] A";
                QueryStr01 += "             INNER JOIN";
                QueryStr01 += "             [@PS_MM050L] B";
                QueryStr01 += "                 ON A.DocEntry = B.DocEntry";
                QueryStr01 += " WHERE       B.U_PODocNum = '" + pDocEntry + "'";
                QueryStr01 += "             AND B.U_POLinNum = '" + pLineID + "'";
                QueryStr01 += "             AND A.Status = 'O'";

                oRecordSet01.DoQuery(QueryStr01);

                if (oRecordSet01.Fields.Item("Cnt").Value > 0)
                {
                    returnValue = false; //삭제 불가
                }
                else
                {
                    returnValue = true; //삭제 가능
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
        /// PS_MM030_TotalAmount_Calculate
        /// </summary>
        private void PS_MM030_TotalAmount_Calculate()
        {
            int i;
            int SumQty = 0;
            double DocTotal = 0;
            double SumWeight = 0;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount -1; i++)
                {
                    DocTotal += Convert.ToDouble(oMat01.Columns.Item("LinTotal").Cells.Item(i + 1).Specific.Value);
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                    {
                    }
                    else
                    {
                        SumQty += Convert.ToInt32(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                    }
                    SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                }
                oForm.Items.Item("DocTotal").Specific.Value = DocTotal;
                oForm.Items.Item("SumQty").Specific.Value = SumQty;
                oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

                oMat01.LoadFromDataSource();
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
        }

        /// <summary>
        /// PS_MM030_TotalAmount_Compare
        /// </summary>
        /// <returns></returns>
        private bool PS_MM030_TotalAmount_Compare()
        {
            bool returnValue = false;
            double DocTotal;
            double EAmount1;
            double EAmount2;
            double EAmount3;
            double MinAmount1 = 0; //비교후 최소값 저장용
            double MinAmount2;
            string errMessage = string.Empty;

            try
            {
                DocTotal = Convert.ToDouble(oForm.Items.Item("DocTotal").Specific.Value);
                EAmount1 = Convert.ToDouble(string.IsNullOrEmpty(oForm.Items.Item("EAmount1").Specific.Value));
                EAmount2 = Convert.ToDouble(string.IsNullOrEmpty(oForm.Items.Item("EAmount2").Specific.Value));
                EAmount3 = Convert.ToDouble(string.IsNullOrEmpty(oForm.Items.Item("EAmount3").Specific.Value));

                if (EAmount1 + EAmount2 + EAmount3 > 0)
                {
                    if (EAmount1 != 0 && EAmount2 != 0)
                    {
                        if (EAmount1 < EAmount2)
                        {
                            MinAmount1 = EAmount1;
                        }
                        else
                        {
                            MinAmount1 = EAmount2;
                        }
                    }
                    else if (EAmount1 == 0 && EAmount2 != 0)
                    {
                        MinAmount1 = EAmount2;
                    }
                    else if (EAmount1 != 0 && EAmount2 == 0)
                    {
                        MinAmount1 = EAmount1;
                    }
                    else if (EAmount1 == 0 && EAmount2 == 0)
                    {
                        MinAmount1 = EAmount3;
                    }
                    if (EAmount3 != 0)
                    {
                        if (MinAmount1 < EAmount3)
                        {
                            MinAmount2 = MinAmount1;
                        }
                        else
                        {
                            MinAmount2 = EAmount3;
                        }
                    }
                    else
                    {
                        MinAmount2 = MinAmount1;
                    }
                }
                else
                {
                    MinAmount2 = DocTotal;
                }
                if (DocTotal <= MinAmount2)
                {
                    returnValue = true;
                }
                else
                {
                    returnValue = false;
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return returnValue;
        }

        // 차후 메일 전송에 대한 이슈가 있을 경우 해당사항을 개발진행해야함.
        ///// <summary>
        ///// PS_MM030_SendMail 
        ///// </summary>
        ///// <returns></returns>
        //private string PS_MM030_SendMail()
        //{
        //    string returnValue = string.Empty;
        //    string errMessage = string.Empty;900
        //    Microsoft.Office.Interop.Outlook.Application objOutlook = default(Microsoft.Office.Interop.Outlook.Application);
        //    Microsoft.Office.Interop.Outlook.MailItem objMail = default(Microsoft.Office.Interop.Outlook.MailItem);

        //    string strToAddress = null;
        //    string strSubject = null;
        //    string strBody = null;

        //    objOutlook = new Microsoft.Office.Interop.Outlook.Application();
        //    objMail = OutlookApplication_definst.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

        //    try
        //    {
        //        strToAddress = MDC_GetData.Get_ReData("ISNULL(E_Mail, '')", "CardCode", "OCRD", "'" + oForm.Items.Item("CardCode").Specific.Value + "'");
        //        //BP마스터에 E-Mail 주소가 없는 경우
        //        if (string.IsNullOrEmpty(strToAddress))
        //        {
        //            PS_MM030_SendMail = "EmailStringEmpty";
        //            throw new Exception();

        //            goto PS_MM030_SendMail_EMailStringEmpty;
        //        }

        //        //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        strSubject = "풍산홀딩스 주문서 첨부 (주문번호:" + oForm.Items.Item("DocNum").Specific.Value + ")";
        //        strBody = "주문서를 첨부합니다." + Strings.Chr(13) + "첨부파일을 참조하십시오.";

        //        objMail.To = strToAddress;
        //        objMail.Subject = strSubject;
        //        objMail.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain;
        //        objMail.Body = strBody;
        //        //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        objMail.Attachments.Add("C:\\ReportExport\\PS_MM030_주문서_" + oForm.Items.Item("DocNum").Specific.Value + ".pdf");

        //        objMail.Send();
        //    }
        //    catch (Exception ex)
        //    {
        //        if (errMessage != string.Empty)
        //        {
        //            PSH_Globals.SBO_Application.MessageBox(errMessage);
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
        //        }
        //    }
        //    finally
        //    {

        //    }

        //    return returnValue;
        //}

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
            int i;
            string DocNumber;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM030_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            
                            }

                            if (PS_MM030_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            //대비견적금액 비교
                            if (PS_MM030_TotalAmount_Compare() == false)
                            {
                                if (PSH_Globals.SBO_Application.MessageBox("대비견적금액이 견적금액보다 적습니다. 정말로 저장하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
                                {
                                }
                                else
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (PS_MM030_oPurchaseOrders_Add() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if(oForm.Items.Item("POFinish").Specific.Selected.Value == "N")
                                {
                                    //행삭제를 한 경우
                                    if (g_deletedCount > 0)
                                    {
                                        for (i = 0; i < itemInfoList.Count; i++)
                                        {
                                            PS_MM030_oPurchaseOrders_LineDelete(Convert.ToInt32(Convert.ToString(itemInfoList[i].LineNum).ToString().Trim()));
                                        }
                                        //삭제행자료 초기화_S
                                        for (i = 0; i < itemInfoList.Count; i++)
                                        {
                                            itemInfoList[i].LineNum = 0;
                                        }
                                        itemInfoList.Clear();//삭제행자료 초기화_E
                                        g_deletedCount = 0; //삭제행 카운트 초기화
                                    }
                                    else
                                    {
                                        PSH_Globals.oCompany.StartTransaction();

                                        if (PS_MM030_oPurchaseOrders_Update() == false)
                                        {
                                            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                            BubbleEvent = false;
                                            return;
                                        }
                                        else
                                        {
                                            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                        }
                                    }
                                }
                            }
                        }
                        PS_MM030_Delete_EmptyRow();
                        oLast_Mode = oForm.Mode;
                        g_deletedCount = 0;
                    }
                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    {
                        oLast_Mode = oForm.Mode;
                    }
                    else if (pVal.ItemUID == "Btn03")
                    {
                        PS_MM030_Copy_Price();
                    }
                    else if (pVal.ItemUID == "Btn04")
                    {
                        DocNumber = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim().Substring(0, 4) + oForm.Items.Item("DocNum").Specific.Value.ToString().Trim();
                        PS_MM035 tempForm = new PS_MM035();
                        tempForm.LoadForm(DocNumber);
                        BubbleEvent = false;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                PS_MM030_Add_MatrixRow(oMat01.RowCount, false);
                            }
                            else if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            {
                                PS_MM030_Add_MatrixRow(oMat01.RowCount, false);
                                PS_MM030_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1291");
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_MM030_Print_Report02);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_MM030_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
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
            string errMessage = string.Empty;

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "CardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "CntcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "PQDocNum")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("PQDocNum").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    if (!string.IsNullOrEmpty(oDS_PS_MM030H.GetValue("U_PODocNum", 0).ToString().Trim()))
                                    {
                                        errMessage = "문서내 구매오더번호가 등록된 상태에서 추가는 불가능합니다.";
                                        throw new Exception();
                                    }
                                    else if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("POType").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("Purchase").Specific.Value.ToString().Trim()))
                                    {
                                        errMessage = "사업장, 품의형태 또는 구매방식을 먼저 선택하세요.";
                                        BubbleEvent = false;
                                        return;
                                    }
                                    else
                                    {
                                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                        BubbleEvent = false;
                                    }
                                }
                            }
                            else if (pVal.ColUID == "ItemCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PS_SM010 tempForm = new PS_SM010();
                                    tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "WhsCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("WhsCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
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
        }

        /// <summary>
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Purchase" || pVal.ItemUID == "BPLId")
                    {
                        oMat01.Clear();
                        oDS_PS_MM030L.Clear();
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_MM030_Add_MatrixRow(0, false);
                        }

                        if (oForm.Items.Item("Purchase").Specific.Value.ToString().Trim() == "30" || oForm.Items.Item("Purchase").Specific.Value.ToString().Trim() == "40" || oForm.Items.Item("Purchase").Specific.Value.ToString().Trim() == "60")
                        {
                            oMat01.Columns.Item("ItemName").Editable = true;
                            oMat01.Columns.Item("OutSize").Editable = true;
                            oMat01.Columns.Item("OutUnit").Editable = true;
                            oMat01.Columns.Item("WhsCode").Editable = false;
                        }
                        else
                        {
                            oMat01.Columns.Item("ItemName").Editable = false;
                            oMat01.Columns.Item("OutSize").Editable = false;
                            oMat01.Columns.Item("OutUnit").Editable = false;
                            oMat01.Columns.Item("WhsCode").Editable = true;
                        }
                    }
                    else if (pVal.ItemUID == "DocCur")
                    {
                        if (oDS_PS_MM030H.GetValue("U_DocCur", 0).ToString().Trim() == "KRW")
                        {
                            oDS_PS_MM030H.SetValue("U_DocRate", 0, Convert.ToString(1));
                        }
                        else
                        {
                            oDS_PS_MM030H.SetValue("U_DocRate", 0, dataHelpClass.Get_ReData("Rate", "Currency", "ORTT", "'" + oDS_PS_MM030H.GetValue("U_DocCur", 0).ToString().Trim() + "'", " AND RateDate = '" + oDS_PS_MM030H.GetValue("U_DocDate", 0).ToString().Trim() + "'"));
                        }
                        if (pVal.ItemChanged == true)
                        {
                            PS_MM030_FlushToItemValue(pVal.ItemUID, 0, "");
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "PQDocNum")
                        {
                            PS_MM010 PS_MM010 = new PS_MM010();
                            PS_MM010.LoadForm(oMat01.Columns.Item("PQDocNum").Cells.Item(pVal.Row).Specific.Value);
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
            int Qty;
            string ItemCode;
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
                        if (pVal.ItemUID == "CardCode")
                        {
                            PS_MM030_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "CntcCode")
                        {
                            PS_MM030_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "DocRate")
                        {
                            PS_MM030_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "PQDocNum")
                            {
                                PS_MM030_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "ItemCode")
                            {
                                PS_MM030_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "Qty")
                            {
                                oMat01.FlushToDataSource();
                                ItemCode = oDS_PS_MM030L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim();
                                Qty = Convert.ToInt32(oDS_PS_MM030L.GetValue("U_Qty", pVal.Row - 1));

                                Calculate_Weight = dataHelpClass.Calculate_Weight(ItemCode, Qty, oForm.Items.Item("BPLId").Specific.Value.ToString().Trim());

                                oDS_PS_MM030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Calculate_Weight));
                                oMat01.LoadFromDataSource();

                                PS_MM030_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "Price")
                            {
                                PS_MM030_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "LinTotal")
                            {
                                PS_MM030_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "WhsCode")
                            {
                                PS_MM030_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
                    PS_MM030_TotalAmount_Calculate();
                    oMat01.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM030H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM030L);
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
            string sQry;
            string DocNum;
            string LineNum;
            string errMessage = string.Empty;
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
                            // 마감일자 Check
                            //if (dataHelpClass.Check_Finish_Status(oDS_PS_MM030H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_MM030H.GetValue("U_DocDate", 0).ToString().Trim().Substring(0, 6)) == false)
                            //{
                            //    errMessage = "마감상태가 잠금입니다. 해당 일자로 취소할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.";
                            //    BubbleEvent = false;
                            //    throw new Exception();
                            //}
                            for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                            {
                                LineNum = oMat01.Columns.Item("LineId").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                                DocNum = oForm.Items.Item("DocNum").Specific.Value.ToString().Trim();

                                if (PS_MM030_LineDelete_Possible(DocNum, LineNum) == false)
                                {
                                    errMessage = "" + i + 1 + "번 라인의 품의(발주)가 가입고등록 되었습니다. 취소할 수 없습니다.";
                                    BubbleEvent = false;
                                    throw new Exception();
                                }
                            }

                            if (PS_MM030_oPurchaseOrders_Cancel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                            }
                            break;
                        case "1286": //닫기
                            // 마감일자 Check
                            //if (dataHelpClass.Check_Finish_Status(oDS_PS_MM030H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_MM030H.GetValue("U_DocDate", 0).ToString().Trim().Substring(0, 6)) == false)
                            //{
                            //    errMessage = "마감상태가 잠금입니다. 해당 일자로 닫기할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.";
                            //    throw new Exception();
                            //}
                            // 부산사업장만 해당되도록 품의취소시 가입고가 남아있을경우 닫기 처리 안됨.
                            if (Convert.ToDouble(oForm.Items.Item("BPLId").Specific.Value) == 2)
                            {
                                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                {
                                    LineNum = oMat01.Columns.Item("LineId").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                                    DocNum = oForm.Items.Item("DocNum").Specific.Value.ToString().Trim();
                                    if (PS_MM030_LineDelete_Possible(DocNum, LineNum) == false)
                                    {
                                        errMessage = "" + i + 1 + "번 라인의 품의(발주)가 가입고등록 되었습니다. 취소할 수 없습니다.";
                                        BubbleEvent = false;
                                        throw new Exception();
                                    }
                                }
                                if (PS_MM030_oPurchaseOrders_Close() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            break;
                        case "1293": //행삭제
                            if (PS_MM030_LineDelete_Possible(oForm.Items.Item("DocNum").Specific.Value, oMat01.Columns.Item("LineId").Cells.Item(oLastColRow01).Specific.Value) == false)
                            {
                                errMessage = "해당 품의는 가입고가 등록되어 행삭제 할 수 없습니다.";
                                BubbleEvent = false;
                                throw new Exception();
                            }
                            else
                            {
                                //구매오더 행삭제를 위해 삭제된 행번호 저장
                                if(g_deletedCount < 1)
                                {
                                    itemInfo.LineNum = Convert.ToInt32(oMat01.Columns.Item("LineNum").Cells.Item(oLastColRow01).Specific.Value);
                                    itemInfoList.Add(itemInfo);
                                    g_deletedCount += 1;
                                }
                                else
                                {
                                    errMessage = "갱신 후 행삭제하세요.";
                                    BubbleEvent = false;
                                    throw new Exception();
                                }
                            }
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
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
                            PS_MM030_FormItemEnabled();
                            oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                                oDS_PS_MM030L.RemoveRecord(oDS_PS_MM030L.Size - 1);
                                // Mat01에 마지막라인(빈라인) 삭제
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();

                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("PQDocNum").Cells.Item(oMat01.RowCount).Specific.Value))
                                {
                                    PS_MM030_Add_MatrixRow(oMat01.RowCount, false);
                                }

                                //문서번호 15647 이전은 최종구매정보를 보여준다.
                                if (Convert.ToInt32(oForm.Items.Item("DocNum").Specific.Value) < 15647)
                                {
                                    //최종구매정보
                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                    {
                                        sQry = "EXEC PS_MM030_03 '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "', '" + oMat01.Columns.Item("ItemCode").Cells.Item(i + 1).Specific.Value.ToString().Trim() + "', '" + oMat01.Columns.Item("ProcCode").Cells.Item(i + 1).Specific.Value.ToString().Trim() + "', '" + oForm.Items.Item("DocNum").Specific.Value.ToString().Trim() + "', '" + oForm.Items.Item("DocDate").Specific.Value.ToString().Trim() + "'";
                                        oRecordSet01.DoQuery(sQry);

                                        oMat01.Columns.Item("LCardCd").Cells.Item(i + 1).Specific.Value = oRecordSet01.Fields.Item("CardCode").Value.ToString().Trim();
                                        oMat01.Columns.Item("LCardNm").Cells.Item(i + 1).Specific.Value = oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim();
                                        if (oRecordSet01.Fields.Item("U_DocDate").Value.ToString("yyyyMMdd").Trim() < "19000102")
                                        {
                                            oMat01.Columns.Item("LDocDate").Cells.Item(i + 1).Specific.Value = "";
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("LDocDate").Cells.Item(i + 1).Specific.Value = oRecordSet01.Fields.Item("U_DocDate").Value.ToString("yyyyMMdd").Trim();
                                        }
                                        oMat01.Columns.Item("LPrice").Cells.Item(i + 1).Specific.Value = oRecordSet01.Fields.Item("U_Price").Value.ToString().Trim();
                                    }
                                }
                                PS_MM030_TotalAmount_Calculate(); //전체금액 계산
                            }
                            break;
                        case "1281": //찾기
                            PS_MM030_FormItemEnabled();
                            oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("SumWeight").Specific.Value = 0;

                            if (string.IsNullOrEmpty(oDocNum))//아이디별 사업장 세팅
                            {
                                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //아이디별 사번 세팅

                                if (dataHelpClass.User_SuperUserYN() == "N")  //수퍼유저인 경우는 사번 미표기(2016.01.08 송명규)
                                {
                                    oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                                }
                            }
                            break;
                        case "1282": //추가
                            oDS_PS_MM030H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            oDS_PS_MM030H.SetValue("U_DueDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            oDS_PS_MM030H.SetValue("U_POStatus", 0, "N");
                            oDS_PS_MM030H.SetValue("U_POFinish", 0, "N");
                            oDS_PS_MM030H.SetValue("U_PurType", 0, "0");
                            oDS_PS_MM030H.SetValue("U_Payment", 0, "10");

                            PS_MM030_Add_MatrixRow(0, true);
                            oForm.Items.Item("SumWeight").Specific.Value = 0;
                            oDS_PS_MM030H.SetValue("U_Payment", 0, "20");
                            //지불조건 기본값을 현금으로 변경(2014.07.23 송명규, 류석균 요청)

                            PS_MM030_Initialization();
                            PS_MM030_FormItemEnabled();
                            PS_MM030_FormClear();
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            PS_MM030_FormItemEnabled();
                            oMat01.AutoResizeColumns();
                            if (oMat01.VisualRowCount > 0)
                            {
                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(oMat01.VisualRowCount).Specific.Value))
                                {
                                    if (oDS_PS_MM030H.GetValue("Status", 0) == "O")
                                    {
                                        PS_MM030_Add_MatrixRow(oMat01.RowCount, false);
                                    }
                                }
                            }
                            break;
                        case "1287": //복제
                            break;
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
        }
    }
}

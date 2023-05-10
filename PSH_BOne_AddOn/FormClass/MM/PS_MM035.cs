using System;
using SAPbouiCOM;
using System.Collections.Generic;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 물품수급계약서등록
    /// </summary>
    internal class PS_MM035 : PSH_BaseClass
    {
        private string oFormUniqueID01;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM035.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PS_MM035_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PS_MM035");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_MM035_CreateItems();
                PS_MM035_DataFind(oFormDocEntry);
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
                oForm.ActiveItem = "DocEntry"; //최초 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_MM035_CreateItems()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                // 사업장
                oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }

                // 문서번호
                oForm.DataSources.UserDataSources.Add("DocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("DocEntry").Specific.DataBind.SetBound(true, "", "DocEntry");

                // 계약일자
                oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
               
                // 계약구분
                oForm.DataSources.UserDataSources.Add("Div", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Div").Specific.ValidValues.Add("", "선택");
                oForm.Items.Item("Div").Specific.ValidValues.Add("1", "장비");
                oForm.Items.Item("Div").Specific.ValidValues.Add("2", "제작,수리");
                oForm.Items.Item("Div").Specific.DataBind.SetBound(true, "", "Div");

                // 거래처코드
                oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

                // 거래처명
                oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

                // 계약명
                oForm.DataSources.UserDataSources.Add("ContName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ContName").Specific.DataBind.SetBound(true, "", "ContName");

                // 계약수량
                oForm.DataSources.UserDataSources.Add("ContQty", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ContQty").Specific.DataBind.SetBound(true, "", "ContQty");

                // 계약기간
                oForm.DataSources.UserDataSources.Add("ContTerm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ContTerm").Specific.DataBind.SetBound(true, "", "ContTerm");

                // 계약금액
                oForm.DataSources.UserDataSources.Add("ContAmt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ContAmt").Specific.DataBind.SetBound(true, "", "ContAmt");

                // 계약이행보증금
                oForm.DataSources.UserDataSources.Add("ContDepo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("ContDepo").Specific.DataBind.SetBound(true, "", "ContDepo");

                // 선급금보증금
                oForm.DataSources.UserDataSources.Add("AdvnDepo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("AdvnDepo").Specific.DataBind.SetBound(true, "", "AdvnDepo");

                // 하자이행보증금
                oForm.DataSources.UserDataSources.Add("FalDepo1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("FalDepo1").Specific.DataBind.SetBound(true, "", "FalDepo1");
                oForm.DataSources.UserDataSources.Add("FalDepo2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("FalDepo2").Specific.DataBind.SetBound(true, "", "FalDepo2");

                // 대금지불
                oForm.DataSources.UserDataSources.Add("PricePay1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("PricePay1").Specific.DataBind.SetBound(true, "", "PricePay1");

                oForm.DataSources.UserDataSources.Add("PricePay2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("PricePay2").Specific.DataBind.SetBound(true, "", "PricePay2");

                oForm.DataSources.UserDataSources.Add("PricePay3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("PricePay3").Specific.DataBind.SetBound(true, "", "PricePay3");

                oForm.DataSources.UserDataSources.Add("PricePay4", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("PricePay4").Specific.DataBind.SetBound(true, "", "PricePay4");

                oForm.DataSources.UserDataSources.Add("PricePay5", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("PricePay5").Specific.DataBind.SetBound(true, "", "PricePay5");

                oForm.DataSources.UserDataSources.Add("PricePay6", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("PricePay6").Specific.DataBind.SetBound(true, "", "PricePay6");

                // 첨부
                oForm.DataSources.UserDataSources.Add("Comment", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 85);
                oForm.Items.Item("Comment").Specific.DataBind.SetBound(true, "", "Comment");

                // 특기사항
                oForm.DataSources.UserDataSources.Add("Dscr1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 85);
                oForm.Items.Item("Dscr1").Specific.DataBind.SetBound(true, "", "Dscr1");

                oForm.DataSources.UserDataSources.Add("Dscr2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 85);
                oForm.Items.Item("Dscr2").Specific.DataBind.SetBound(true, "", "Dscr2");

                oForm.DataSources.UserDataSources.Add("Dscr3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 85);
                oForm.Items.Item("Dscr3").Specific.DataBind.SetBound(true, "", "Dscr3");

                oForm.DataSources.UserDataSources.Add("Dscr4", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 85);
                oForm.Items.Item("Dscr4").Specific.DataBind.SetBound(true, "", "Dscr4");

                oForm.DataSources.UserDataSources.Add("Dscr5", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 85);
                oForm.Items.Item("Dscr5").Specific.DataBind.SetBound(true, "", "Dscr5");

                oForm.DataSources.UserDataSources.Add("Dscr6", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 85);
                oForm.Items.Item("Dscr6").Specific.DataBind.SetBound(true, "", "Dscr6");

                oForm.DataSources.UserDataSources.Add("Dscr7", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 85);
                oForm.Items.Item("Dscr7").Specific.DataBind.SetBound(true, "", "Dscr7");

                oForm.DataSources.UserDataSources.Add("Dscr8", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 85);
                oForm.Items.Item("Dscr8").Specific.DataBind.SetBound(true, "", "Dscr8");

                oForm.DataSources.UserDataSources.Add("Dscr9", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 85);
                oForm.Items.Item("Dscr9").Specific.DataBind.SetBound(true, "", "Dscr9");

                oForm.DataSources.UserDataSources.Add("Dscr10", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 85);
                oForm.Items.Item("Dscr10").Specific.DataBind.SetBound(true, "", "Dscr10");

                // 계약문구
                oForm.DataSources.UserDataSources.Add("ContText", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("ContText").Specific.ValidValues.Add("", "선택");
                oForm.Items.Item("ContText").Specific.ValidValues.Add("1", "'구매자'와'공급자'는 계약일반 조건을 첨부하여 물품수급 계약을 체결하고, 그 증거로 전자계약서를 작성하여 쌍방이 공인인증서로 서명후 보관한다.");
                oForm.Items.Item("ContText").Specific.ValidValues.Add("2", "'구매자'와'공급자'는 물품수급 계약을 체결하고, 그 증거로 전자계약서를 작성하여 쌍방이 공인인증서로 서명후 보관한다.");
                oForm.Items.Item("ContText").Specific.ValidValues.Add("3", "'구매자'와'공급자'는 계약일반 조건을 첨부하여 물품수급 계약을 체결하고, 그 증거로 계약서 2부를 작성 쌍방이 날인후 각각1부씩 보관한다.");
                oForm.Items.Item("ContText").Specific.DataBind.SetBound(true, "", "ContText");

                // 공공업체 대표자직함 성명
                oForm.DataSources.UserDataSources.Add("GName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("GName").Specific.DataBind.SetBound(true, "", "GName");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// PS_MM035_DataFind
        /// </summary>
        /// <param name="DocNumber"></param>
        private void PS_MM035_DataFind(string DocNumber)
        {
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (string.IsNullOrEmpty(DocNumber))
                {
                    errMessage = "문서번호가 없습니다. 확인바랍니다.";
                    throw new Exception();
                }

                sQry = " Select * From [PS_MM035] Where DocEntry = '" + DocNumber + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    PS_MM035_DataNew(DocNumber);
                }
                else
                {
                    oForm.Items.Item("FStatus").Specific.Caption = "화면상태 : 조회";
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;

                    oForm.Items.Item("BPLId").Specific.Select(oRecordSet.Fields.Item("BPLId").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("DocEntry").Specific.Value = oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim();
                    oForm.Items.Item("DocDate").Specific.Value = oRecordSet.Fields.Item("DocDate").Value.ToString().Trim();
                    oForm.Items.Item("Div").Specific.Select(oRecordSet.Fields.Item("Div").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("CardCode").Specific.Value = oRecordSet.Fields.Item("CardCode").Value.ToString().Trim();
                    oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item("CardName").Value.ToString().Trim();
                    oForm.Items.Item("ContName").Specific.Value = oRecordSet.Fields.Item("ContName").Value.ToString().Trim();
                    oForm.Items.Item("ContQty").Specific.Value = oRecordSet.Fields.Item("ContQty").Value.ToString().Trim();
                    oForm.Items.Item("ContTerm").Specific.Value = oRecordSet.Fields.Item("ContTerm").Value.ToString().Trim();
                    oForm.Items.Item("ContAmt").Specific.Value = oRecordSet.Fields.Item("ContAmt").Value.ToString().Trim();
                    oForm.Items.Item("ContDepo").Specific.Value = oRecordSet.Fields.Item("ContDepo").Value.ToString().Trim();
                    oForm.Items.Item("AdvnDepo").Specific.Value = oRecordSet.Fields.Item("AdvnDepo").Value.ToString().Trim();
                    oForm.Items.Item("FalDepo1").Specific.Value = oRecordSet.Fields.Item("FalDepo1").Value.ToString().Trim();
                    oForm.Items.Item("FalDepo2").Specific.Value = oRecordSet.Fields.Item("FalDepo2").Value.ToString().Trim();
                    oForm.Items.Item("PricePay1").Specific.Value = oRecordSet.Fields.Item("PricePay1").Value.ToString().Trim();
                    oForm.Items.Item("PricePay2").Specific.Value = oRecordSet.Fields.Item("PricePay2").Value.ToString().Trim();
                    oForm.Items.Item("PricePay3").Specific.Value = oRecordSet.Fields.Item("PricePay3").Value.ToString().Trim();
                    oForm.Items.Item("PricePay4").Specific.Value = oRecordSet.Fields.Item("PricePay4").Value.ToString().Trim();
                    oForm.Items.Item("PricePay5").Specific.Value = oRecordSet.Fields.Item("PricePay5").Value.ToString().Trim();
                    oForm.Items.Item("PricePay6").Specific.Value = oRecordSet.Fields.Item("PricePay6").Value.ToString().Trim();
                    oForm.Items.Item("Comment").Specific.Value = oRecordSet.Fields.Item("Comment").Value.ToString().Trim();
                    oForm.Items.Item("Dscr1").Specific.Value = oRecordSet.Fields.Item("Dscr1").Value.ToString().Trim();
                    oForm.Items.Item("Dscr2").Specific.Value = oRecordSet.Fields.Item("Dscr2").Value.ToString().Trim();
                    oForm.Items.Item("Dscr3").Specific.Value = oRecordSet.Fields.Item("Dscr3").Value.ToString().Trim();
                    oForm.Items.Item("Dscr4").Specific.Value = oRecordSet.Fields.Item("Dscr4").Value.ToString().Trim();
                    oForm.Items.Item("Dscr5").Specific.Value = oRecordSet.Fields.Item("Dscr5").Value.ToString().Trim();
                    oForm.Items.Item("Dscr6").Specific.Value = oRecordSet.Fields.Item("Dscr6").Value.ToString().Trim();
                    oForm.Items.Item("Dscr7").Specific.Value = oRecordSet.Fields.Item("Dscr7").Value.ToString().Trim();
                    oForm.Items.Item("Dscr8").Specific.Value = oRecordSet.Fields.Item("Dscr8").Value.ToString().Trim();
                    oForm.Items.Item("Dscr9").Specific.Value = oRecordSet.Fields.Item("Dscr9").Value.ToString().Trim();
                    oForm.Items.Item("Dscr10").Specific.Value = oRecordSet.Fields.Item("Dscr10").Value.ToString().Trim();
                    oForm.Items.Item("ContText").Specific.Select(oRecordSet.Fields.Item("ContText").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("GName").Specific.Value = oRecordSet.Fields.Item("GName").Value.ToString().Trim();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_MM035_DataNew
        /// </summary>
        /// <param name="DocNumber"></param>
        private void PS_MM035_DataNew(string DocNumber)
        {
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                sQry =  " EXEC PS_MM035_01 '" + DocNumber + "'";
                oRecordSet.DoQuery(sQry);
               
                if (oRecordSet.RecordCount > 0)
                {
                    oForm.Items.Item("FStatus").Specific.Caption = "화면상태 : 등록";
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;

                    oForm.Items.Item("BPLId").Specific.Select(oRecordSet.Fields.Item("BPLId").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("DocEntry").Specific.Value = oRecordSet.Fields.Item("DocEntry").Value.Trim();
                    oForm.Items.Item("DocDate").Specific.Value = oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd").Trim();
                    oForm.Items.Item("Div").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("CardCode").Specific.Value = oRecordSet.Fields.Item("CardCode").Value.Trim();
                    oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item("CardName").Value.Trim();
                    oForm.Items.Item("ContName").Specific.Value = oRecordSet.Fields.Item("ContName").Value.Trim();
                    oForm.Items.Item("ContQty").Specific.Value = oRecordSet.Fields.Item("ContQty").Value.Trim();
                    oForm.Items.Item("ContTerm").Specific.Value = oRecordSet.Fields.Item("ContTerm").Value.Trim();
                    oForm.Items.Item("ContAmt").Specific.Value = "일금" + oRecordSet.Fields.Item("ContAmt_Han").Value.Trim() + "원정" +
                                                                 "(￦"  + Convert.ToString(oRecordSet.Fields.Item("ContAmt").Value) +
                                                                 "원, VAT 별도)";
                    oForm.Items.Item("ContDepo").Specific.Value = "";
                    oForm.Items.Item("AdvnDepo").Specific.Value = "";
                    oForm.Items.Item("FalDepo1").Specific.Value = "";
                    oForm.Items.Item("FalDepo2").Specific.Value = "";
                    oForm.Items.Item("PricePay1").Specific.Value = "";
                    oForm.Items.Item("PricePay2").Specific.Value = "";
                    oForm.Items.Item("PricePay3").Specific.Value = "";
                    oForm.Items.Item("PricePay4").Specific.Value = "";
                    oForm.Items.Item("PricePay5").Specific.Value = "";
                    oForm.Items.Item("PricePay6").Specific.Value = "";
                    oForm.Items.Item("Comment").Specific.Value = "";
                    oForm.Items.Item("Dscr1").Specific.Value = "";
                    oForm.Items.Item("Dscr2").Specific.Value = "";
                    oForm.Items.Item("Dscr3").Specific.Value = "";
                    oForm.Items.Item("Dscr4").Specific.Value = "";
                    oForm.Items.Item("Dscr5").Specific.Value = "";
                    oForm.Items.Item("Dscr6").Specific.Value = "";
                    oForm.Items.Item("Dscr7").Specific.Value = "";
                    oForm.Items.Item("Dscr8").Specific.Value = "";
                    oForm.Items.Item("Dscr9").Specific.Value = "";
                    oForm.Items.Item("Dscr10").Specific.Value = "";
                    oForm.Items.Item("ContText").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("GName").Specific.Value = oRecordSet.Fields.Item("GName").Value.Trim();
                }
                else
                {
                    errMessage = "문서번호를 생성할 수 없습니다. 확인바랍니다.";
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_MM035_SAVE
        /// </summary>
        private void PS_MM035_SAVE()
        {
            string BPLId;
            string DocEntry;
            string DocDate;
            string Div;
            string CardCode;
            string CardName;
            string ContName;
            string ContQty;
            string ContTerm;
            string ContAmt;
            string ContDepo;
            string AdvnDepo;
            string FalDepo1;
            string FalDepo2;
            string PricePay1;
            string PricePay2;
            string PricePay3;
            string PricePay4;
            string PricePay5;
            string PricePay6;
            string Comment;
            string Dscr1;
            string Dscr2;
            string Dscr3;
            string Dscr4;
            string Dscr5;
            string Dscr6;
            string Dscr7;
            string Dscr8;
            string Dscr9;
            string Dscr10;
            string ContText;
            string GName;
            string errMessage = string.Empty;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                BPLId     = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                DocEntry  = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                DocDate   = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                Div       = oForm.Items.Item("Div").Specific.Value.ToString().Trim();
                CardCode  = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                CardName  = oForm.Items.Item("CardName").Specific.Value.ToString().Trim();
                ContName  = oForm.Items.Item("ContName").Specific.Value.ToString().Trim();
                ContQty   = oForm.Items.Item("ContQty").Specific.Value.ToString().Trim();
                ContTerm  = oForm.Items.Item("ContTerm").Specific.Value.ToString().Trim();
                ContAmt   = oForm.Items.Item("ContAmt").Specific.Value.ToString().Trim();
                ContDepo  = oForm.Items.Item("ContDepo").Specific.Value.ToString().Trim();
                AdvnDepo  = oForm.Items.Item("AdvnDepo").Specific.Value.ToString().Trim();
                FalDepo1  = oForm.Items.Item("FalDepo1").Specific.Value.ToString().Trim();
                FalDepo2  = oForm.Items.Item("FalDepo2").Specific.Value.ToString().Trim();
                PricePay1 = oForm.Items.Item("PricePay1").Specific.Value.ToString().Trim();
                PricePay2 = oForm.Items.Item("PricePay2").Specific.Value.ToString().Trim();
                PricePay3 = oForm.Items.Item("PricePay3").Specific.Value.ToString().Trim();
                PricePay4 = oForm.Items.Item("PricePay4").Specific.Value.ToString().Trim();
                PricePay5 = oForm.Items.Item("PricePay5").Specific.Value.ToString().Trim();
                PricePay6 = oForm.Items.Item("PricePay6").Specific.Value.ToString().Trim();
                Comment   = oForm.Items.Item("Comment").Specific.Value.ToString().Trim();
                Dscr1     = oForm.Items.Item("Dscr1").Specific.Value.ToString().Trim();
                Dscr2     = oForm.Items.Item("Dscr2").Specific.Value.ToString().Trim();
                Dscr3     = oForm.Items.Item("Dscr3").Specific.Value.ToString().Trim();
                Dscr4     = oForm.Items.Item("Dscr4").Specific.Value.ToString().Trim();
                Dscr5     = oForm.Items.Item("Dscr5").Specific.Value.ToString().Trim();
                Dscr6     = oForm.Items.Item("Dscr6").Specific.Value.ToString().Trim();
                Dscr7     = oForm.Items.Item("Dscr7").Specific.Value.ToString().Trim();
                Dscr8     = oForm.Items.Item("Dscr8").Specific.Value.ToString().Trim();
                Dscr9     = oForm.Items.Item("Dscr9").Specific.Value.ToString().Trim();
                Dscr10    = oForm.Items.Item("Dscr10").Specific.Value.ToString().Trim();
                ContText  = oForm.Items.Item("ContText").Specific.Value.ToString().Trim();
                GName     = oForm.Items.Item("GName").Specific.Value.ToString().Trim();

                if (string.IsNullOrWhiteSpace(BPLId))
                {
                    errMessage = "사업장코드를 확인 하세요";
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(DocEntry))
                {
                    errMessage = "문서번호를 확인 하세요";
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(Div))
                {
                    errMessage = "계약구분을 확인 하세요";
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(DocDate))
                {
                    errMessage = "계약일자를 확인 하세요";
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(CardCode))
                {
                    errMessage = "거래처코드를 확인 하세요";
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(PricePay1) || string.IsNullOrWhiteSpace(PricePay3) ||
                    string.IsNullOrWhiteSpace(PricePay5) || string.IsNullOrWhiteSpace(PricePay6))
                {
                    errMessage = "대금지불조건을 확인 하세요 ";
                    throw new Exception();
                }
                if (Convert.ToDouble(PricePay1) + Convert.ToDouble(PricePay3) + Convert.ToDouble(PricePay5) + Convert.ToDouble(PricePay6) != 100)
                {
                    errMessage = "대금지불 전체조건의 합이100% 이여야 합니다.";
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(ContText))
                {
                    errMessage = "계약문구선택을 확인 하세요";
                    throw new Exception();
                }

                sQry = " Select * From [PS_MM035] Where DocEntry = '" + DocEntry + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount <= 0)  // INSERT
                {
                    sQry = "INSERT INTO [PS_MM035]";
                    sQry += " (";
                    sQry += "BPLId,";
                    sQry += "DocEntry,";
                    sQry += "DocDate,";
                    sQry += "Div,";
                    sQry += "CardCode,";
                    sQry += "CardName,";
                    sQry += "ContName,";
                    sQry += "ContQty,";
                    sQry += "ContTerm,";
                    sQry += "ContAmt,";
                    sQry += "ContDepo,";
                    sQry += "AdvnDepo,";
                    sQry += "FalDepo1,";
                    sQry += "FalDepo2,";
                    sQry += "PricePay1,";
                    sQry += "PricePay2,";
                    sQry += "PricePay3,";
                    sQry += "PricePay4,";
                    sQry += "PricePay5,";
                    sQry += "PricePay6,";
                    sQry += "Comment,";
                    sQry += "Dscr1,";
                    sQry += "Dscr2,";
                    sQry += "Dscr3,";
                    sQry += "Dscr4,";
                    sQry += "Dscr5,";
                    sQry += "Dscr6,";
                    sQry += "Dscr7,";
                    sQry += "Dscr8,";
                    sQry += "Dscr9,";
                    sQry += "Dscr10,";
                    sQry += "ContText,";
                    sQry += "GName";
                    sQry += " ) ";
                    sQry += "VALUES(";
                    sQry += "'" + BPLId + "',";
                    sQry += "'" + DocEntry + "',";
                    sQry += "'" + DocDate + "',";
                    sQry += "'" + Div + "',";
                    sQry += "'" + CardCode + "',";
                    sQry += "'" + CardName + "',";
                    sQry += "'" + ContName + "',";
                    sQry += "'" + ContQty + "',";
                    sQry += "'" + ContTerm + "',";
                    sQry += "'" + ContAmt + "',";
                    sQry += "'" + ContDepo + "',";
                    sQry += "'" + AdvnDepo + "',";
                    sQry += "'" + FalDepo1 + "',";
                    sQry += "'" + FalDepo2 + "',";
                    sQry += "'" + PricePay1 + "',";
                    sQry += "'" + PricePay2 + "',";
                    sQry += "'" + PricePay3 + "',";
                    sQry += "'" + PricePay4 + "',";
                    sQry += "'" + PricePay5 + "',";
                    sQry += "'" + PricePay6 + "',";
                    sQry += "'" + Comment + "',";
                    sQry += "'" + Dscr1 + "',";
                    sQry += "'" + Dscr2 + "',";
                    sQry += "'" + Dscr3 + "',";
                    sQry += "'" + Dscr4 + "',";
                    sQry += "'" + Dscr5 + "',";
                    sQry += "'" + Dscr6 + "',";
                    sQry += "'" + Dscr7 + "',";
                    sQry += "'" + Dscr8 + "',";
                    sQry += "'" + Dscr9 + "',";
                    sQry += "'" + Dscr10 + "',";
                    sQry += "'" + ContText + "',";
                    sQry += "'" + GName + "'";
                    sQry += " ) ";
                    oRecordSet.DoQuery(sQry);
                    oForm.Items.Item("FStatus").Specific.Caption = "화면상태 : 등록완료!";
                }
                else // UPDATE
                {
                    sQry = "UPDATE [PS_MM035] SET ";
                    sQry += "DocDate   = '" + DocDate   + "',";
                    sQry += "Div       = '" + Div       + "',";
                    sQry += "CardCode  = '" + CardCode  + "',";
                    sQry += "CardName  = '" + CardName  + "',";
                    sQry += "ContName  = '" + ContName  + "',";
                    sQry += "ContQty   = '" + ContQty   + "',";
                    sQry += "ContTerm  = '" + ContTerm  + "',";
                    sQry += "ContAmt   = '" + ContAmt   + "',";
                    sQry += "ContDepo  = '" + ContDepo  + "',";
                    sQry += "AdvnDepo  = '" + AdvnDepo  + "',";
                    sQry += "FalDepo1  = '" + FalDepo1  + "',";
                    sQry += "FalDepo2  = '" + FalDepo2  + "',";
                    sQry += "PricePay1 = '" + PricePay1 + "',";
                    sQry += "PricePay2 = '" + PricePay2 + "',";
                    sQry += "PricePay3 = '" + PricePay3 + "',";
                    sQry += "PricePay4 = '" + PricePay4 + "',";
                    sQry += "PricePay5 = '" + PricePay5 + "',";
                    sQry += "PricePay6 = '" + PricePay6 + "',";
                    sQry += "Comment   = '" + Comment   + "',";
                    sQry += "Dscr1     = '" + Dscr1     + "',";
                    sQry += "Dscr2     = '" + Dscr2     + "',";
                    sQry += "Dscr3     = '" + Dscr3     + "',";
                    sQry += "Dscr4     = '" + Dscr4     + "',";
                    sQry += "Dscr5     = '" + Dscr5     + "',";
                    sQry += "Dscr6     = '" + Dscr6     + "',";
                    sQry += "Dscr7     = '" + Dscr7     + "',";
                    sQry += "Dscr8     = '" + Dscr8     + "',";
                    sQry += "Dscr9     = '" + Dscr9     + "',";
                    sQry += "Dscr10    = '" + Dscr10    + "',";
                    sQry += "ContText  = '" + ContText  + "',";
                    sQry += "GName     = '" + GName     + "'";
                    sQry += " Where BPLId = '" + BPLId + "' And DocEntry = '" + DocEntry + "'";
                    oRecordSet.DoQuery(sQry);
                    oForm.Items.Item("FStatus").Specific.Caption = "화면상태 : 수정완료!";
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
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_MM035_Delete 데이타 삭제
        /// </summary>
        private void PS_MM035_Delete()
        {
            string BPLId;
            string DocEntry;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", 2, "예", "아니오") == 1)
                {
                    sQry = " Delete From [PS_MM035] Where BPLId = '" + BPLId + "' And DocEntry = '" + DocEntry + "'";
                    oRecordSet.DoQuery(sQry);
                    PS_MM035_Clear();
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
        /// PS_MM035_Clear
        /// </summary>
        private void PS_MM035_Clear()
        {
            try
            {
                oForm.Freeze(true);
                oForm.Items.Item("FStatus").Specific.Caption = "";
                oForm.Items.Item("BPLId").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("DocEntry").Specific.Value = "";
                oForm.Items.Item("DocDate").Specific.Value = "";
                oForm.Items.Item("Div").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("CardCode").Specific.Value = "";
                oForm.Items.Item("CardName").Specific.Value = "";
                oForm.Items.Item("ContName").Specific.Value = "";
                oForm.Items.Item("ContQty").Specific.Value = "";
                oForm.Items.Item("ContTerm").Specific.Value = "";
                oForm.Items.Item("ContAmt").Specific.Value = "";
                oForm.Items.Item("ContDepo").Specific.Value = "";
                oForm.Items.Item("AdvnDepo").Specific.Value = "";
                oForm.Items.Item("FalDepo1").Specific.Value = "";
                oForm.Items.Item("FalDepo2").Specific.Value = "";
                oForm.Items.Item("PricePay1").Specific.Value = "";
                oForm.Items.Item("PricePay2").Specific.Value = "";
                oForm.Items.Item("PricePay3").Specific.Value = "";
                oForm.Items.Item("PricePay4").Specific.Value = "";
                oForm.Items.Item("PricePay5").Specific.Value = "";
                oForm.Items.Item("PricePay6").Specific.Value = "";
                oForm.Items.Item("Comment").Specific.Value = "";
                oForm.Items.Item("Dscr1").Specific.Value = "";
                oForm.Items.Item("Dscr2").Specific.Value = "";
                oForm.Items.Item("Dscr3").Specific.Value = "";
                oForm.Items.Item("Dscr4").Specific.Value = "";
                oForm.Items.Item("Dscr5").Specific.Value = "";
                oForm.Items.Item("Dscr6").Specific.Value = "";
                oForm.Items.Item("Dscr7").Specific.Value = "";
                oForm.Items.Item("Dscr8").Specific.Value = "";
                oForm.Items.Item("Dscr9").Specific.Value = "";
                oForm.Items.Item("Dscr10").Specific.Value = "";
                oForm.Items.Item("ContText").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("GName").Specific.Value = "";
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
        /// PS_MM035_Print_Report
        /// </summary>
        [STAThread]
        private void PS_MM035_Print_Report()
        {
            string WinTitle;
            string ReportName;
            string BPLId;
            string DocEntry;
            string DocNumber;

            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            
            try
            {
                BPLId = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
                DocNumber = oForm.Items.Item("DocEntry").Specific.Value.Trim();
                DocEntry = DocNumber.Substring(4);

                WinTitle = "[PS_MM035] 물품수급계약서 출력";
                ReportName = "PS_MM035_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                //Formula

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocNumber));

                //SubReport Parameter
                dataPackSubReportParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry, "PS_MM035_01_SUB1"));

                formHelpClass.OpenCrystalReport(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
            string sQry;
            string Result;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn01")  // 저장
                    {
                        sQry = "select U_POFinish from [@PS_MM030H] where DocEntry ='" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim().Substring(4) + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result == "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("종결된 품의(발부)이므로 등록 수정할 수 없습니다.");
                        }
                        else
                        {
                            PS_MM035_SAVE();
                        }
                    }
                    else if (pVal.ItemUID == "Btn_del")  // 삭제
                    {
                        sQry = "select U_POFinish from [@PS_MM030H] where DocEntry ='" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim().Substring(4) + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result == "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("종결된 품의(발주)이므로 삭제할 수 없습니다.");
                        }
                        else
                        {
                            PS_MM035_Delete();
                        }
                    }
                    else if (pVal.ItemUID == "Btn_Prt")  // 출력
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_MM035_Print_Report);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }

                    
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "CardCode":
                                sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value;
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                        case "1283":
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1293":
                            break;
                        case "1281":
                            break;
                        case "1282":
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            break;
                        case "1282": //문서추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                        case "1293": // 행삭제
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }
}

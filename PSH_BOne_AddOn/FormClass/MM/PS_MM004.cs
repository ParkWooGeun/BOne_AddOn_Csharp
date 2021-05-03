using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 운송확인증등록(출력)
    /// </summary>
    internal class PS_MM004 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_MM004H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM004L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// 클래스내에서 공통으로 사용되는 폼 호출 메소드
        /// </summary>
        private new void LoadForm()
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM004.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM004_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM004");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_MM004_CreateItems();
                PS_MM004_SetForm();
                PS_MM004_SetComboBox();
                PS_MM004_LoadCaption();

                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1286", false); // 닫기
                oForm.EnableMenu("1287", false); // 복제
                oForm.EnableMenu("1285", false); // 복원
                oForm.EnableMenu("1284", false); // 취소
                oForm.EnableMenu("1293", false); // 행삭제
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", true);

                oMat01.Columns.Item("Check").Visible = false; //Matrix선택 체크박스 Visible = False
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
        /// Form 호출(메인메뉴)
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            try
            {
                LoadForm();
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
        /// Form 호출(외부Form)
        /// </summary>
        /// <param name="oFormName"></param>
        /// <param name="oFormDocEntry"></param>
        public void LoadForm(string oFormName, string oFormDocEntry)
        {
            string sQry;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                LoadForm();

                oForm.Freeze(true);

                if (!string.IsNullOrEmpty(oFormDocEntry))
                {
                    sQry = "Select Cnt = Count(*) From [@PS_MM004H] Where U_ObjType = '" + oFormName + "' And U_DocNo = '" + oFormDocEntry + "'";
                    RecordSet01.DoQuery(sQry);

                    if (RecordSet01.Fields.Item(0).Value > 0)
                    {
                        sQry = "Select DocEntry, U_DocDate From [@PS_MM004H] Where U_ObjType = '" + oFormName + "' And U_DocNo = '" + oFormDocEntry + "'";
                        RecordSet01.DoQuery(sQry);

                        oForm.Items.Item("SDocDateF").Specific.Value = RecordSet01.Fields.Item(1).Value.ToString("yyyyMMdd");
                        oForm.Items.Item("SDocDateT").Specific.Value = RecordSet01.Fields.Item(1).Value.ToString("yyyyMMdd");

                        oForm.Items.Item("SObjType").Specific.Value = oFormName;
                        oForm.Items.Item("SDocNo").Specific.Value = oFormDocEntry;

                        oForm.Items.Item("BtnSearch").Click();

                        PS_MM004_RET(RecordSet01.Fields.Item(0).Value.ToString().Trim());
                    }
                    else
                    {
                        oForm.Items.Item("ObjType").Specific.Value = oFormName;
                        oForm.Items.Item("DocNo").Specific.Value = oFormDocEntry;

                        oForm.Items.Item("SectionF").Specific.Value = "창원";

                        if (oFormName == "PS_SD040")
                        {
                            sQry = "Select Distinct b.U_ItmBsort From [@PS_SD040L] a Inner join OITM b On a.U_ItemCode = b.ItemCode and a.DocEntry = '" + oFormDocEntry + "'";
                            RecordSet01.DoQuery(sQry);

                            if (RecordSet01.RecordCount > 0)
                            {
                                oForm.Items.Item("ItmBsort").Specific.Value = RecordSet01.Fields.Item(0).Value.ToString().Trim();
                                if (RecordSet01.Fields.Item(0).Value.ToString().Trim() == "102")
                                {
                                    oForm.Items.Item("SectionT").Specific.Value = "안강";
                                }
                            }

                            sQry = "Select U_TranCard, U_TranCode, U_Tonnage, U_Destin, U_DocDate From [@PS_SD040H] Where DocEntry = '" + oFormDocEntry + "'";
                            RecordSet01.DoQuery(sQry);

                            if (RecordSet01.RecordCount > 0)
                            {
                                oForm.Items.Item("TranCard").Specific.Value = RecordSet01.Fields.Item(0).Value.ToString().Trim();
                                oForm.Items.Item("TranCode").Specific.Value = RecordSet01.Fields.Item(1).Value.ToString().Trim();
                                oForm.Items.Item("Tonnage").Specific.Value = RecordSet01.Fields.Item(2).Value.ToString().Trim();
                                oForm.Items.Item("Destin").Specific.Value = RecordSet01.Fields.Item(3).Value.ToString().Trim();

                                oForm.Items.Item("DocDate").Specific.Value = RecordSet01.Fields.Item(4).Value.ToString("yyyyMMdd");
                                oForm.Items.Item("SDocDateF").Specific.Value = RecordSet01.Fields.Item(4).Value.ToString("yyyyMMdd");
                            }
                            oForm.Items.Item("LocCode").Specific.Value = "07";
                        }
                        else if (oFormName == "PS_PP095")
                        {
                            sQry = "Select U_Destin, U_Section, U_Tonnage, U_TranCard, U_TranCode, U_TWeight, U_DocDate From [@PS_PP095H] Where DocEntry = '" + oFormDocEntry + "'";
                            RecordSet01.DoQuery(sQry);

                            if (RecordSet01.RecordCount > 0)
                            {
                                oForm.Items.Item("Destin").Specific.Value = RecordSet01.Fields.Item(0).Value.ToString().Trim();
                                oForm.Items.Item("SectionT").Specific.Value = RecordSet01.Fields.Item(1).Value.ToString().Trim();
                                oForm.Items.Item("Tonnage").Specific.Value = RecordSet01.Fields.Item(2).Value.ToString().Trim();
                                oForm.Items.Item("TranCard").Specific.Value = RecordSet01.Fields.Item(3).Value.ToString().Trim();
                                oForm.Items.Item("TranCode").Specific.Value = RecordSet01.Fields.Item(4).Value.ToString().Trim();
                                oForm.Items.Item("TWeight").Specific.Value = RecordSet01.Fields.Item(5).Value.ToString().Trim();
                                oForm.Items.Item("DocDate").Specific.Value = RecordSet01.Fields.Item(6).Value.ToString("yyyyMMdd");
                                oForm.Items.Item("SDocDateF").Specific.Value = RecordSet01.Fields.Item(6).Value.ToString("yyyyMMdd");
                            }

                            oForm.Items.Item("ItmBsort").Specific.Value = "104";

                            sQry = "Select DocDate = Convert(Char(8),Max(U_DocDate),112) From [@PS_MM003H] Where U_BPLId = '" + oForm.Items.Item("BPLId").Specific.Value + "'";
                            sQry = sQry + " And U_DocDate <= '" + oForm.Items.Item("DocDate").Specific.Value + "'";
                            RecordSet01.DoQuery(sQry);

                            sQry = "  Select    b.U_LocCode";
                            sQry += " From      [@PS_MM003H] a";
                            sQry += "           Inner Join";
                            sQry += "           [@PS_MM003L] b";
                            sQry += "               On a.Code = b.Code ";
                            sQry += " Where     a.U_BPLId = '" + oForm.Items.Item("BPLId").Specific.Value + "'";
                            sQry += "           And a.U_Docdate = '" + RecordSet01.Fields.Item(0).Value.ToString().Trim() + "'";
                            sQry += "           And b.U_LocName Like '%" + oForm.Items.Item("SectionT").Specific.Value + "%'";
                            RecordSet01.DoQuery(sQry);

                            oForm.Items.Item("LocCode").Specific.Value = RecordSet01.Fields.Item(0).Value.ToString().Trim();

                            //실중량
                            sQry = "Select Weight = Sum(U_Weight) From [@PS_PP095L] Where DocEntry = '" + oFormDocEntry + "'";
                            RecordSet01.DoQuery(sQry);
                            if (RecordSet01.RecordCount > 0)
                            {
                                oForm.Items.Item("Weight").Specific.Value = RecordSet01.Fields.Item(0).Value.ToString().Trim();
                            }
                        }
                    }
                }
                else
                {
                    oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
                    oForm.Items.Item("SDocDateF").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
                }

                oForm.Items.Item("SDocDateT").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
                oForm.Items.Item("DocDate").Click();
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
        /// 화면 Item 생성
        /// </summary>
        private void PS_MM004_CreateItems()
        {
            try
            {
                oDS_PS_MM004H = oForm.DataSources.DBDataSources.Item("@PS_MM004H");
                oDS_PS_MM004L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("SDocDateF", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("SDocDateF").Specific.DataBind.SetBound(true, "", "SDocDateF");

                oForm.DataSources.UserDataSources.Add("SDocDateT", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("SDocDateT").Specific.DataBind.SetBound(true, "", "SDocDateT");

                oForm.DataSources.UserDataSources.Add("SBPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("SBPLId").Specific.DataBind.SetBound(true, "", "SBPLId");

                oForm.DataSources.UserDataSources.Add("SItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SItmBsort").Specific.DataBind.SetBound(true, "", "SItmBsort");

                oForm.DataSources.UserDataSources.Add("SLocCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SLocCode").Specific.DataBind.SetBound(true, "", "SLocCode");

                oForm.DataSources.UserDataSources.Add("SWay", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SWay").Specific.DataBind.SetBound(true, "", "SWay");

                oForm.DataSources.UserDataSources.Add("STranCard", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("STranCard").Specific.DataBind.SetBound(true, "", "STranCard");

                oForm.DataSources.UserDataSources.Add("STonnage", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("STonnage").Specific.DataBind.SetBound(true, "", "STonnage");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면 초기화
        /// </summary>
        private void PS_MM004_SetForm()
        {
            string sQry;
            string docEntry;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_MM004H]";
                RecordSet01.DoQuery(sQry);

                docEntry = RecordSet01.Fields.Item(0).Value.ToString().Trim();

                if (string.IsNullOrEmpty(docEntry) || docEntry == "0")
                {
                    oDS_PS_MM004H.SetValue("DocEntry", 0, "1");
                }
                else
                {
                    oDS_PS_MM004H.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(docEntry) + 1));
                }
                
                //==========기준정보==========
                oDS_PS_MM004H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd")); //일자
                oDS_PS_MM004H.SetValue("U_ItmBsort", 0, ""); //품목분류
                oDS_PS_MM004H.SetValue("U_ItmBName", 0, ""); //품목분류명
                oDS_PS_MM004H.SetValue("U_ObjType", 0, ""); //참조Object
                oDS_PS_MM004H.SetValue("U_DocNo", 0, ""); //참조문서번호
                oDS_PS_MM004H.SetValue("U_LocCode", 0, ""); //운송지역코드
                oDS_PS_MM004H.SetValue("U_LocName", 0, ""); //운송지역명
                oDS_PS_MM004H.SetValue("U_Way", 0, "10"); //편도/회로

                oDS_PS_MM004H.SetValue("U_SectionF", 0, ""); //운행구간1
                oDS_PS_MM004H.SetValue("U_SectionT", 0, ""); //운행구간2
                oDS_PS_MM004H.SetValue("U_Destin", 0, ""); //도착장소
                oDS_PS_MM004H.SetValue("U_TranCard", 0, ""); //운송업체
                oDS_PS_MM004H.SetValue("U_TranCode", 0, ""); //차량번호
                oDS_PS_MM004H.SetValue("U_Tonnage", 0, ""); //차종(톤수)
                oDS_PS_MM004H.SetValue("U_ItemName", 0, ""); //품명
                oDS_PS_MM004H.SetValue("U_Weight", 0, ""); //운송량
                oDS_PS_MM004H.SetValue("U_TWeight", 0, ""); //계근량
                oDS_PS_MM004H.SetValue("U_PassYN", 0, "N"); //경유Y/N
                oDS_PS_MM004H.SetValue("U_Amt", 0, ""); //기본운임
                oDS_PS_MM004H.SetValue("U_PassAmt", 0, ""); //경유금액
                oDS_PS_MM004H.SetValue("U_TotAmt", 0, ""); //총금액
                oDS_PS_MM004H.SetValue("U_Comments", 0, ""); //비고

                //==========기준정보==========
                oForm.Items.Item("SLocCode").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue); //운송지역
                oForm.Items.Item("SWay").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue); //편도/회로
                oForm.Items.Item("STonnage").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue); //차종(톤수)
                oForm.Items.Item("SObjType").Specific.Value = "";
                oForm.Items.Item("SDocNo").Specific.Value = "";

                oForm.Items.Item("ItmBsort").Click();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }

        /// <summary>
        /// 콤보박스 세팅
        /// </summary>
        private void PS_MM004_SetComboBox()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
                oForm.Items.Item("SBPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                if (dataHelpClass.User_BPLID() == "1")
                {
                    oForm.Items.Item("SectionF").Specific.Value = "창원";
                }

                oForm.Items.Item("SItmBsort").Specific.ValidValues.Add("", "");
                dataHelpClass.Set_ComboList(oForm.Items.Item("SItmBsort").Specific, "select Code, Name from [@PSH_ITMBSORT] Where U_PudYN = 'Y' order by Code", "101", false, false);
                oForm.Items.Item("SItmBsort").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.Items.Item("Way").Specific.ValidValues.Add("10", "편도");
                oForm.Items.Item("Way").Specific.ValidValues.Add("20", "회로");
                oForm.Items.Item("Way").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.Items.Item("PassYN").Specific.ValidValues.Add("N", "일반");
                oForm.Items.Item("PassYN").Specific.ValidValues.Add("Y", "경유지");
                oForm.Items.Item("PassYN").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.Items.Item("SLocCode").Specific.ValidValues.Add("", "");
                dataHelpClass.Set_ComboList(oForm.Items.Item("SLocCode").Specific, "select Distinct b.U_LocCode, b.U_LocName from [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code Where a.U_BPLId = '" + dataHelpClass.User_BPLID() + "' order by b.U_LocCode", "", false, false);
                oForm.Items.Item("SLocCode").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.Items.Item("STonnage").Specific.ValidValues.Add("", "");
                dataHelpClass.Set_ComboList(oForm.Items.Item("STonnage").Specific, "select Distinct b.U_Tonnage, b.U_TonNm from [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code Where a.U_BPLId = '" + dataHelpClass.User_BPLID() + "' order by b.U_Tonnage", "", false, false);
                oForm.Items.Item("STonnage").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.Items.Item("SWay").Specific.ValidValues.Add("", "");
                oForm.Items.Item("SWay").Specific.ValidValues.Add("10", "편도");
                oForm.Items.Item("SWay").Specific.ValidValues.Add("20", "회로");

                oMat01.Columns.Item("Way").ValidValues.Add("10", "편도");
                oMat01.Columns.Item("Way").ValidValues.Add("20", "회로");

                sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("BPLId").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 버튼 title 변경(추가/확인/갱신)
        /// </summary>
        private void PS_MM004_LoadCaption()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
                    oForm.Items.Item("BtnDelete").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
                    oForm.Items.Item("BtnDelete").Enabled = true;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 기 정보 조회
        /// </summary>
        /// <param name="oDocEntry"></param>
        private void PS_MM004_RET(string oDocEntry)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                sQry = "  SELECT    DocEntry,";
                sQry += "           U_ObjType,";
                sQry += "           U_DocNo,";
                sQry += "           U_BPLId,";
                sQry += "           U_DocDate,";
                sQry += "           U_ItmBsort,";
                sQry += "           U_ItmBName,";
                sQry += "           U_LocCode,";
                sQry +="            U_LocName,";
                sQry += "           U_Way,";
                sQry += "           U_SectionF,";
                sQry += "           U_SectionT,";
                sQry += "           U_Destin,";
                sQry += "           U_TranCard,";
                sQry += "           U_TranCode,";
                sQry += "           U_Tonnage,";
                sQry += "           U_ItemName,";
                sQry += "           U_Weight,";
                sQry += "           U_TWeight,";
                sQry += "           U_PassYN,";
                sQry += "           U_Amt,";
                sQry += "           U_PassAmt,";
                sQry += "           U_TotAmt,";
                sQry += "           U_Comments ";
                sQry += " FROM      [@PS_MM004H]";
                sQry += " WHERE     DocEntry = '" + oDocEntry + "'";

                oRecordSet01.DoQuery(sQry);

                //DataSource를 이용하여 각 컨트롤에 값을 출력
                oDS_PS_MM004H.SetValue("DocEntry", 0, oRecordSet01.Fields.Item("DocEntry").Value); //관리번호
                oDS_PS_MM004H.SetValue("U_ObjType", 0, oRecordSet01.Fields.Item("U_ObjType").Value.ToString().Trim()); //참조Object
                oDS_PS_MM004H.SetValue("U_DocNo", 0, oRecordSet01.Fields.Item("U_DocNo").Value); //참조문서번호
                oDS_PS_MM004H.SetValue("U_BPLId", 0, oRecordSet01.Fields.Item("U_BPLId").Value.ToString().Trim()); //사업장
                oDS_PS_MM004H.SetValue("U_DocDate", 0, oRecordSet01.Fields.Item("U_DocDate").Value.ToString("yyyyMMdd")); //일자
                oDS_PS_MM004H.SetValue("U_ItmBsort", 0, oRecordSet01.Fields.Item("U_ItmBsort").Value.ToString().Trim()); //분류코드
                oDS_PS_MM004H.SetValue("U_ItmBName", 0, oRecordSet01.Fields.Item("U_ItmBName").Value.ToString().Trim()); //분류명
                oDS_PS_MM004H.SetValue("U_LocCode", 0, oRecordSet01.Fields.Item("U_LocCode").Value.ToString().Trim()); //운송지역코드
                oDS_PS_MM004H.SetValue("U_LocName", 0, oRecordSet01.Fields.Item("U_LocName").Value.ToString().Trim()); //운송지역명
                oDS_PS_MM004H.SetValue("U_Way", 0, oRecordSet01.Fields.Item("U_Way").Value.ToString().Trim()); //편도/회로
                oDS_PS_MM004H.SetValue("U_SectionF", 0, oRecordSet01.Fields.Item("U_SectionF").Value.ToString().Trim()); //운행구간시작
                oDS_PS_MM004H.SetValue("U_SectionT", 0, oRecordSet01.Fields.Item("U_SectionT").Value.ToString().Trim()); //운행구간종료
                oDS_PS_MM004H.SetValue("U_Destin", 0, oRecordSet01.Fields.Item("U_Destin").Value.ToString().Trim()); //도착장소
                oDS_PS_MM004H.SetValue("U_TranCard", 0, oRecordSet01.Fields.Item("U_TranCard").Value.ToString().Trim()); //운송업체
                oDS_PS_MM004H.SetValue("U_TranCode", 0, oRecordSet01.Fields.Item("U_TranCode").Value.ToString().Trim()); //차량번호
                oDS_PS_MM004H.SetValue("U_Tonnage", 0, oRecordSet01.Fields.Item("U_Tonnage").Value); //차종(톤수)
                oDS_PS_MM004H.SetValue("U_ItemName", 0, oRecordSet01.Fields.Item("U_ItemName").Value); //품명
                oDS_PS_MM004H.SetValue("U_Weight", 0, oRecordSet01.Fields.Item("U_Weight").Value); //운송량
                oDS_PS_MM004H.SetValue("U_TWeight", 0, oRecordSet01.Fields.Item("U_TWeight").Value); //계근중량
                oDS_PS_MM004H.SetValue("U_PassYN", 0, oRecordSet01.Fields.Item("U_PassYN").Value); //경유여부
                oDS_PS_MM004H.SetValue("U_Amt", 0, oRecordSet01.Fields.Item("U_Amt").Value); //기본운임
                oDS_PS_MM004H.SetValue("U_PassAmt", 0, oRecordSet01.Fields.Item("U_PassAmt").Value); //경유지운임
                oDS_PS_MM004H.SetValue("U_TotAmt", 0, oRecordSet01.Fields.Item("U_TotAmt").Value); //총금액
                oDS_PS_MM004H.SetValue("U_Comments", 0, oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim()); //비고

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                PS_MM004_LoadCaption();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메트릭스 행추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_MM004_Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PS_MM004H.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PS_MM004H.Offset = oRow;
                oDS_PS_MM004H.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_MM004_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            double TWeight;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                switch (oUID)
                {
                    case "DocDate":
                        oForm.Items.Item("SDocDateF").Specific.Value = oForm.Items.Item("DocDate").Specific.Value;
                        oForm.Items.Item("SDocDateT").Specific.Value = oForm.Items.Item("DocDate").Specific.Value;
                        break;
                    case "ItmBsort": //품목분류
                        sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" + oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("ItmBName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                        if (!string.IsNullOrEmpty(oForm.Items.Item("ObjType").Specific.Value.ToString().Trim()))
                        {
                            oForm.Items.Item("ItemName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        oForm.Items.Item("SItmBsort").Specific.Select(oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                        break;
                    case "LocCode": //운송지역코드
                        sQry = "  Select    b.U_LocName,";
                        sQry += "           b.U_PassAmt";
                        sQry += " From      [@PS_MM003H] a";
                        sQry += "           Inner Join";
                        sQry += "           [@PS_MM003L] b";
                        sQry += "                ON a.Code = b.Code";
                        sQry += " Where     a.U_BPLId =  '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
                        sQry += "           And b.U_LocCode = '" + oForm.Items.Item("LocCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("LocName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                        if (oForm.Items.Item("PassYN").Specific.Value == "Y")
                        {
                            oForm.Items.Item("PassAmt").Specific.Value = oRecordSet01.Fields.Item(1).Value.ToString().Trim();
                        }
                        else
                        {
                            oForm.Items.Item("PassAmt").Specific.Value = "0";
                        }

                        TWeight = Convert.ToDouble(oForm.Items.Item("TWeight").Specific.Value);
                        
                        if (TWeight > 0) //총계근중량으로 운송비 다시계산
                        {
                            sQry = "  Select    Amt = Isnull(b.U_Amt,0)";
                            sQry += " From      [@PS_MM003H] a";
                            sQry += "           Inner Join";
                            sQry += "           [@PS_MM003L] b";
                            sQry += "               On a.Code = b.Code";
                            sQry += " Where     a.U_BPLId =  '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
                            sQry += "           And b.U_LocCode = '" + oForm.Items.Item("LocCode").Specific.Value.ToString().Trim() + "'";
                            sQry += "           And " + TWeight + " Between b.U_KgF And b.U_KgT ";
                            oRecordSet01.DoQuery(sQry);

                            if (oForm.Items.Item("Way").Specific.Value.ToString().Trim() == "10")
                            {
                                oForm.Items.Item("Amt").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            }
                            else if (oForm.Items.Item("Way").Specific.Value.ToString().Trim() == "20")
                            {
                                oForm.Items.Item("Amt").Specific.Value = Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) * 0.8);
                            }
                        }
                        break;
                    case "Amt":
                    case "PassAmt":
                        oForm.Items.Item("TotAmt").Specific.Value = Convert.ToString(Convert.ToDouble(oForm.Items.Item("Amt").Specific.Value) + Convert.ToDouble(oForm.Items.Item("PassAmt").Specific.Value));
                        break;
                    case "TWeight": //중량입력시 화물운송료 계산
                        TWeight = Convert.ToDouble(oForm.Items.Item("TWeight").Specific.Value); //MM095 총중량 구간으로 금액 Select
                        if (TWeight > 0) //총계근중량으로 운송비 다시계산
                        {
                            sQry = "  Select    Amt = Isnull(b.U_Amt,0)";
                            sQry += " From      [@PS_MM003H] a";
                            sQry += "           Inner Join";
                            sQry += "           [@PS_MM003L] b";
                            sQry += "               On a.Code = b.Code";
                            sQry += " Where     a.U_BPLId =  '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
                            sQry += "           And b.U_LocCode = '" + oForm.Items.Item("LocCode").Specific.Value.ToString().Trim() + "'";
                            sQry += "           And " + TWeight + " Between b.U_KgF And b.U_KgT ";
                            oRecordSet01.DoQuery(sQry);

                            if (oForm.Items.Item("Way").Specific.Value.ToString().Trim() == "10")
                            {
                                oForm.Items.Item("Amt").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            }
                            else if (oForm.Items.Item("Way").Specific.Value.ToString().Trim() == "20")
                            {
                                oForm.Items.Item("Amt").Specific.Value = Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) * 0.8);
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        #region PS_MM004_MTX01
        //		public void PS_MM004_MTX01()
        //		{
        //			//******************************************************************************
        //			//Function ID : PS_MM004_MTX01()
        //			//해당모듈 : PS_MM004
        //			//기능 : 데이터 조회
        //			//인수 : 없음
        //			//반환값 : 없음
        //			//특이사항 : 없음
        //			//******************************************************************************
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			short i = 0;
        //			string sQry = null;
        //			short ErrNum = 0;

        //			SAPbobsCOM.Recordset oRecordSet01 = null;
        //			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			string SBPLID = null;
        //			//사업장
        //			string SDocDateF = null;
        //			//조회기간1
        //			string SDocDateT = null;
        //			//조회기간2
        //			string SItmBsort = null;
        //			//폼목분류
        //			string SLocCode = null;
        //			//운송지역
        //			string SWay = null;
        //			//편도/회로
        //			string STranCard = null;
        //			//운송업체
        //			string STonnage = null;
        //			//차종(톤수)
        //			string SObjType = null;
        //			//참조Object
        //			string SDocNo = null;
        //			//참조문서번호


        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SBPLID = Strings.Trim(oForm.Items.Item("SBPLId").Specific.Value);
        //			//사업장
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SDocDateF = Strings.Trim(oForm.Items.Item("SDocDateF").Specific.Value);
        //			//조회시작일
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SDocDateT = Strings.Trim(oForm.Items.Item("SDocDateT").Specific.Value);
        //			//조회종료일
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SItmBsort = Strings.Trim(oForm.Items.Item("SItmBsort").Specific.Value);
        //			//품목분류
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SLocCode = Strings.Trim(oForm.Items.Item("SLocCode").Specific.Value);
        //			//운송지역
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SWay = Strings.Trim(oForm.Items.Item("SWay").Specific.Value);
        //			//편도/회로
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			STranCard = Strings.Trim(oForm.Items.Item("STranCard").Specific.Value);
        //			//운송업체
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			STonnage = Strings.Trim(oForm.Items.Item("STonnage").Specific.Value);
        //			//차종(톤수)
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SObjType = Strings.Trim(oForm.Items.Item("SObjType").Specific.Value);
        //			//참조Object
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SDocNo = Strings.Trim(oForm.Items.Item("SDocNo").Specific.Value);
        //			//참조문서번호

        //			SAPbouiCOM.ProgressBar ProgBar01 = null;
        //			ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

        //			oForm.Freeze(true);

        //			sQry = "                EXEC [PS_MM004_01] ";
        //			sQry = sQry + "'" + SBPLID + "',";
        //			//사업장
        //			sQry = sQry + "'" + SDocDateF + "',";
        //			//조회시작일
        //			sQry = sQry + "'" + SDocDateT + "',";
        //			//조회종료일
        //			sQry = sQry + "'" + SItmBsort + "',";
        //			//품목분류
        //			sQry = sQry + "'" + SLocCode + "',";
        //			//운송지역
        //			sQry = sQry + "'" + SWay + "',";
        //			//편도/회로
        //			sQry = sQry + "'" + STranCard + "',";
        //			//운송업체
        //			sQry = sQry + "'" + STonnage + "',";
        //			//차종(톤수)
        //			sQry = sQry + "'" + SObjType + "',";
        //			//참조Object
        //			sQry = sQry + "'" + SDocNo + "'";
        //			//참조문서번호

        //			oRecordSet01.DoQuery(sQry);

        //			oMat01.Clear();
        //			oDS_PS_MM004L.Clear();
        //			oMat01.FlushToDataSource();
        //			oMat01.LoadFromDataSource();

        //			if ((oRecordSet01.RecordCount == 0)) {

        //				ErrNum = 1;

        //				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

        //				//        Call PS_MM004_PS_MM004_Add_MatrixRow(0, True)
        //				PS_MM004_LoadCaption();

        //				goto PS_MM004_MTX01_Error;

        //				return;
        //			}

        //			for (i = 0; i <= oRecordSet01.RecordCount - 1; i++) {
        //				if (i + 1 > oDS_PS_MM004L.Size) {
        //					oDS_PS_MM004L.InsertRecord((i));
        //				}

        //				oMat01.AddRow();
        //				oDS_PS_MM004L.Offset = i;

        //				oDS_PS_MM004L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //				oDS_PS_MM004L.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet01.Fields.Item("DocEntry").Value));
        //				//관리번호
        //				oDS_PS_MM004L.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("BPLId").Value));
        //				//사업장
        //				oDS_PS_MM004L.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("DocDate").Value));
        //				//일자
        //				oDS_PS_MM004L.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet01.Fields.Item("ItmBName").Value));
        //				//품목분류명
        //				oDS_PS_MM004L.SetValue("U_ColReg05", i, Strings.Trim(oRecordSet01.Fields.Item("LocName").Value));
        //				//운송지역
        //				oDS_PS_MM004L.SetValue("U_ColReg06", i, Strings.Trim(oRecordSet01.Fields.Item("Way").Value));
        //				//편도/회로
        //				oDS_PS_MM004L.SetValue("U_ColReg07", i, Strings.Trim(oRecordSet01.Fields.Item("SectionF").Value));
        //				//운행구간(출발)
        //				oDS_PS_MM004L.SetValue("U_ColReg08", i, Strings.Trim(oRecordSet01.Fields.Item("SectionT").Value));
        //				//운행구간(도착)
        //				oDS_PS_MM004L.SetValue("U_ColReg09", i, Strings.Trim(oRecordSet01.Fields.Item("Destin").Value));
        //				//도착장소
        //				oDS_PS_MM004L.SetValue("U_ColReg10", i, Strings.Trim(oRecordSet01.Fields.Item("TranCard").Value));
        //				//운송업체
        //				oDS_PS_MM004L.SetValue("U_ColReg11", i, Strings.Trim(oRecordSet01.Fields.Item("TranCode").Value));
        //				//차량번호
        //				oDS_PS_MM004L.SetValue("U_ColReg12", i, Strings.Trim(oRecordSet01.Fields.Item("Tonnage").Value));
        //				//차종(톤수)
        //				oDS_PS_MM004L.SetValue("U_ColReg13", i, Strings.Trim(oRecordSet01.Fields.Item("ItemName").Value));
        //				//품명
        //				oDS_PS_MM004L.SetValue("U_ColNum01", i, Strings.Trim(oRecordSet01.Fields.Item("Weight").Value));
        //				//운송량
        //				oDS_PS_MM004L.SetValue("U_ColSum01", i, Strings.Trim(oRecordSet01.Fields.Item("TotAmt").Value));
        //				//금액
        //				oDS_PS_MM004L.SetValue("U_ColReg15", i, Strings.Trim(oRecordSet01.Fields.Item("PassYN").Value));
        //				//비고
        //				oDS_PS_MM004L.SetValue("U_ColReg16", i, Strings.Trim(oRecordSet01.Fields.Item("Comments").Value));
        //				//비고

        //				oRecordSet01.MoveNext();
        //				ProgBar01.Value = ProgBar01.Value + 1;
        //				ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";

        //			}

        //			oMat01.LoadFromDataSource();
        //			oMat01.AutoResizeColumns();
        //			ProgBar01.Stop();
        //			oForm.Freeze(false);

        //			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			ProgBar01 = null;
        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;
        //			return;
        //			PS_MM004_MTX01_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			//    ProgBar01.Stop
        //			oForm.Freeze(false);
        //			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			ProgBar01 = null;
        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;

        //			if (ErrNum == 1) {
        //				MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.", ref "W");
        //			} else {
        //				MDC_Com.MDC_GF_Message(ref "PS_MM004_MTX01_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //			}
        //		}
        #endregion

        #region PS_MM004_DeleteData
        //		public void PS_MM004_DeleteData()
        //		{
        //			//******************************************************************************
        //			//Function ID : PS_MM004_DeleteData()
        //			//해당모듈 : PS_MM004
        //			//기능 : 기본정보 삭제
        //			//인수 : 없음
        //			//반환값 : 없음
        //			//특이사항 : 없음
        //			//******************************************************************************
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			short i = 0;
        //			string sQry = null;
        //			short ErrNum = 0;

        //			SAPbobsCOM.Recordset oRecordSet01 = null;
        //			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			string DocEntry = null;

        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {

        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);

        //				sQry = "SELECT COUNT(*) FROM [@PS_MM004H] WHERE DocEntry = '" + DocEntry + "'";
        //				oRecordSet01.DoQuery(sQry);

        //				if ((oRecordSet01.RecordCount == 0)) {
        //					ErrNum = 1;
        //					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        //					goto PS_MM004_DeleteData_Error;
        //				} else {
        //					sQry = "EXEC PS_MM004_04 '" + DocEntry + "'";
        //					oRecordSet01.DoQuery(sQry);
        //				}
        //			}

        //			MDC_Com.MDC_GF_Message(ref "삭제 완료!", ref "S");

        //			//    Call PS_MM004_SetForm

        //			//    oForm.Mode = fm_ADD_MODE

        //			//    Call oForm.Items("BtnSearch").Click(ct_Regular)

        //			//    oMat01.Clear
        //			//    oMat01.FlushToDataSource
        //			//    oMat01.LoadFromDataSource
        //			//    Call PS_MM004_PS_MM004_Add_MatrixRow(0, True)

        //			return;
        //			PS_MM004_DeleteData_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;
        //			if (ErrNum == 1) {
        //				MDC_Com.MDC_GF_Message(ref "삭제대상이 없습니다. 확인하세요.", ref "W");
        //			} else {
        //				MDC_Com.MDC_GF_Message(ref "PS_MM004_DeleteData_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //			}
        //		}
        #endregion

        #region PS_MM004_UpdateData
        //		public bool PS_MM004_UpdateData()
        //		{
        //			bool functionReturnValue = false;
        //			//******************************************************************************
        //			//Function ID : PS_MM004_UpdateData()
        //			//해당모듈 : PS_MM004
        //			//기능 : 기본정보를 수정
        //			//인수 : 없음
        //			//반환값 : 없음
        //			//특이사항 : 없음
        //			//******************************************************************************
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			short i = 0;
        //			short j = 0;
        //			string sQry = null;
        //			SAPbobsCOM.Recordset RecordSet01 = null;
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			short DocEntry = 0;
        //			//관리번호
        //			string BPLId = null;
        //			//사업장
        //			string DocDate = null;
        //			//일자
        //			string ItmBsort = null;
        //			//분류코드
        //			string ItmBName = null;
        //			//분류명
        //			string LocCode = null;
        //			//운송지역코드
        //			string LocName = null;
        //			//운송지역명
        //			string Way = null;
        //			//편도/회로
        //			string SectionF = null;
        //			//운행구간출발
        //			string SectionT = null;
        //			//운행구간도착
        //			string Destin = null;
        //			//도착장소
        //			string TranCard = null;
        //			//운송업체
        //			string TranCode = null;
        //			//차량번호
        //			double Tonnage = 0;
        //			//차종(톤수)
        //			string ItemName = null;
        //			//품명
        //			double Weight = 0;
        //			//실중량
        //			double TWeight = 0;
        //			//총중량
        //			string PassYN = null;
        //			double Amt = 0;
        //			//기본운임
        //			double PassAmt = 0;
        //			//경유운임
        //			double TotAmt = 0;
        //			//총운임
        //			string Comments = null;
        //			//비고

        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			DocEntry = Convert.ToInt16(Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value));
        //			//관리번호
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			BPLId = Strings.Trim(oForm.Items.Item("BPLId").Specific.Value);
        //			//사업장
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			DocDate = Strings.Trim(oForm.Items.Item("DocDate").Specific.Value);
        //			//일자
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			ItmBsort = Strings.Trim(oForm.Items.Item("ItmBsort").Specific.Value);
        //			//품목분류코드
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			ItmBName = Strings.Trim(oForm.Items.Item("ItmBName").Specific.Value);
        //			//품목분류명
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			LocCode = Strings.Trim(oForm.Items.Item("LocCode").Specific.Value);
        //			//운송지역코드
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			LocName = Strings.Trim(oForm.Items.Item("LocName").Specific.Value);
        //			//운송지역명
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Way = Strings.Trim(oForm.Items.Item("Way").Specific.Value);
        //			//편도/회로
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SectionF = Strings.Trim(oForm.Items.Item("SectionF").Specific.Value);
        //			//운행구간출발
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SectionT = Strings.Trim(oForm.Items.Item("SectionT").Specific.Value);
        //			//운행구간도착
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Destin = Strings.Trim(oForm.Items.Item("Destin").Specific.Value);
        //			//도착장소
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TranCard = Strings.Trim(oForm.Items.Item("TranCard").Specific.Value);
        //			//운송업체
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TranCode = Strings.Trim(oForm.Items.Item("TranCode").Specific.Value);
        //			//차량번호
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Tonnage = oForm.Items.Item("Tonnage").Specific.Value;
        //			//차종(톤수)
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			ItemName = oForm.Items.Item("ItemName").Specific.Value;
        //			//품명
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Weight = oForm.Items.Item("Weight").Specific.Value;
        //			//실중량
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TWeight = oForm.Items.Item("TWeight").Specific.Value;
        //			//계근중량
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			PassYN = oForm.Items.Item("PassYN").Specific.Value;
        //			//경유지YN
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Amt = oForm.Items.Item("Amt").Specific.Value;
        //			//기본운임
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			PassAmt = oForm.Items.Item("PassAmt").Specific.Value;
        //			//경유지운임
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TotAmt = oForm.Items.Item("TotAmt").Specific.Value;
        //			//총운임
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Comments = Strings.Trim(oForm.Items.Item("Comments").Specific.Value);
        //			//비고

        //			if (string.IsNullOrEmpty(Strings.Trim(Convert.ToString(DocEntry)))) {
        //				MDC_Com.MDC_GF_Message(ref "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", ref "E");
        //				functionReturnValue = false;
        //				return functionReturnValue;
        //			}

        //			sQry = "                EXEC [PS_MM004_03] ";
        //			sQry = sQry + DocEntry + ",";
        //			//관리번호
        //			sQry = sQry + "'" + BPLId + "',";
        //			//사업장
        //			sQry = sQry + "'" + DocDate + "',";
        //			//출장번호1
        //			sQry = sQry + "'" + ItmBsort + "',";
        //			//출장번호2
        //			sQry = sQry + "'" + ItmBName + "',";
        //			//사원번호
        //			sQry = sQry + "'" + LocCode + "',";
        //			//사원성명
        //			sQry = sQry + "'" + LocName + "',";
        //			//출장지
        //			sQry = sQry + "'" + Way + "',";
        //			//출장지상세
        //			sQry = sQry + "'" + SectionF + "',";
        //			//작번
        //			sQry = sQry + "'" + SectionT + "',";
        //			//시작일자
        //			sQry = sQry + "'" + Destin + "',";
        //			//시작시각
        //			sQry = sQry + "'" + TranCard + "',";
        //			//종료일자
        //			sQry = sQry + "'" + TranCode + "',";
        //			//종료시각
        //			sQry = sQry + Tonnage + ",";
        //			//목적
        //			sQry = sQry + "'" + ItemName + "',";
        //			//품명
        //			sQry = sQry + Weight + ",";
        //			//실중량
        //			sQry = sQry + TWeight + ",";
        //			//총중량
        //			sQry = sQry + "'" + PassYN + "',";
        //			//경유지YN
        //			sQry = sQry + Amt + ",";
        //			//기본운임
        //			sQry = sQry + PassAmt + ",";
        //			//경유지운임
        //			sQry = sQry + TotAmt + ",";
        //			//총운임
        //			sQry = sQry + "'" + Comments + "'";
        //			//비고

        //			RecordSet01.DoQuery(sQry);

        //			MDC_Com.MDC_GF_Message(ref "수정 완료!", ref "S");

        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			functionReturnValue = true;
        //			return functionReturnValue;
        //			PS_MM004_UpdateData_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			MDC_Com.MDC_GF_Message(ref "PS_MM004_UpdateData_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //			return functionReturnValue;
        //		}
        #endregion

        #region PS_MM004_AddData
        //		public bool PS_MM004_AddData()
        //		{
        //			bool functionReturnValue = false;
        //			//******************************************************************************
        //			//Function ID : PS_MM004_AddData()
        //			//해당모듈 : PS_MM004
        //			//기능 : 데이터 INSERT
        //			//인수 : 없음
        //			//반환값 : 성공여부
        //			//특이사항 : 없음
        //			//******************************************************************************
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			short i = 0;
        //			string sQry = null;

        //			SAPbobsCOM.Recordset RecordSet01 = null;
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			SAPbobsCOM.Recordset RecordSet02 = null;
        //			RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			int DocEntry = 0;
        //			string ObjType = null;
        //			//참조문서타입
        //			int DocNo = 0;
        //			string BPLId = null;
        //			//사업장
        //			string DocDate = null;
        //			//일자
        //			string ItmBsort = null;
        //			//분류코드
        //			string ItmBName = null;
        //			//분류명
        //			string LocCode = null;
        //			//운송지역코드
        //			string LocName = null;
        //			//운송지역명
        //			string Way = null;
        //			//편도/회로
        //			string SectionF = null;
        //			//운행구간출발
        //			string SectionT = null;
        //			//운행구간도착
        //			string Destin = null;
        //			//도착장소
        //			string TranCard = null;
        //			//운송업체
        //			string TranCode = null;
        //			//차량번호
        //			double Tonnage = 0;
        //			//차종(톤수)
        //			string ItemName = null;
        //			//품명

        //			double Weight = 0;
        //			//실중량
        //			double TWeight = 0;
        //			//총중량

        //			string PassYN = null;
        //			//경유지 YN
        //			double Amt = 0;
        //			//금액
        //			double PassAmt = 0;
        //			//경유지금액
        //			double TotAmt = 0;
        //			//총금액
        //			string Comments = null;
        //			//비고

        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			ObjType = Strings.Trim(oForm.Items.Item("ObjType").Specific.Value);
        //			//참조문서Object
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			DocNo = Conversion.Val(oForm.Items.Item("DocNo").Specific.Value);
        //			//참조문서번호
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			BPLId = Strings.Trim(oForm.Items.Item("BPLId").Specific.Value);
        //			//사업장
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			DocDate = Strings.Trim(oForm.Items.Item("DocDate").Specific.Value);
        //			//일자
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			ItmBsort = Strings.Trim(oForm.Items.Item("ItmBsort").Specific.Value);
        //			//품목분류
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			ItmBName = Strings.Trim(oForm.Items.Item("ItmBName").Specific.Value);
        //			//품목분류명
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			LocCode = Strings.Trim(oForm.Items.Item("LocCode").Specific.Value);
        //			//운송지역
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			LocName = Strings.Trim(oForm.Items.Item("LocName").Specific.Value);
        //			//운송지역명
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Way = Strings.Trim(oForm.Items.Item("Way").Specific.Value);
        //			//편도/회로
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SectionF = Strings.Trim(oForm.Items.Item("SectionF").Specific.Value);
        //			//운행구간(출발)
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SectionT = Strings.Trim(oForm.Items.Item("SectionT").Specific.Value);
        //			//운행구간(도착)
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Destin = Strings.Trim(oForm.Items.Item("Destin").Specific.Value);
        //			//도착장소
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TranCard = Strings.Trim(oForm.Items.Item("TranCard").Specific.Value);
        //			//운송업체
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TranCode = Strings.Trim(oForm.Items.Item("TranCode").Specific.Value);
        //			//차량번호
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Tonnage = oForm.Items.Item("Tonnage").Specific.Value;
        //			//차종(톤수)
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			ItemName = oForm.Items.Item("ItemName").Specific.Value;
        //			//품명

        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Weight = oForm.Items.Item("Weight").Specific.Value;
        //			//실중량
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TWeight = oForm.Items.Item("TWeight").Specific.Value;
        //			//계근중량
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			PassYN = oForm.Items.Item("PassYN").Specific.Value;
        //			//경유지YN
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Amt = oForm.Items.Item("Amt").Specific.Value;
        //			//금액
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			PassAmt = oForm.Items.Item("PassAmt").Specific.Value;
        //			//경유지금액
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			TotAmt = oForm.Items.Item("TotAmt").Specific.Value;
        //			//총금액
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Comments = Strings.Trim(oForm.Items.Item("Comments").Specific.Value);
        //			//비고

        //			//DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
        //			sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_MM004H]";
        //			RecordSet01.DoQuery(sQry);

        //			if (Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) == 0) {
        //				DocEntry = 1;
        //			} else {
        //				DocEntry = Convert.ToDouble(Strings.Trim(RecordSet01.Fields.Item(0).Value)) + 1;
        //			}

        //			sQry = "                EXEC [PS_MM004_02] ";
        //			sQry = sQry + DocEntry + ",";
        //			//관리번호
        //			sQry = sQry + "'" + ObjType + "',";
        //			//참조문서Object
        //			sQry = sQry + DocNo + ",";
        //			//참조문서번호
        //			sQry = sQry + "'" + BPLId + "',";
        //			//사업장
        //			sQry = sQry + "'" + DocDate + "',";
        //			//일자
        //			sQry = sQry + "'" + ItmBsort + "',";
        //			//분류코드
        //			sQry = sQry + "'" + ItmBName + "',";
        //			//분류명
        //			sQry = sQry + "'" + LocCode + "',";
        //			//운송지역
        //			sQry = sQry + "'" + LocName + "',";
        //			//운송지역명
        //			sQry = sQry + "'" + Way + "',";
        //			//편도/회로
        //			sQry = sQry + "'" + SectionF + "',";
        //			//운행구간출발
        //			sQry = sQry + "'" + SectionT + "',";
        //			//운행구간도착
        //			sQry = sQry + "'" + Destin + "',";
        //			//도착장소
        //			sQry = sQry + "'" + TranCard + "',";
        //			//운송업체
        //			sQry = sQry + "'" + TranCode + "',";
        //			//차량번호
        //			sQry = sQry + Tonnage + ",";
        //			//차종(톤수)
        //			sQry = sQry + "'" + ItemName + "',";
        //			//품명
        //			sQry = sQry + Weight + ",";
        //			//실중량
        //			sQry = sQry + TWeight + ",";
        //			//총중량
        //			sQry = sQry + "'" + PassYN + "',";
        //			//경유지YN
        //			sQry = sQry + Amt + ",";
        //			//금액
        //			sQry = sQry + PassAmt + ",";
        //			//경유지금액
        //			sQry = sQry + TotAmt + ",";
        //			//금액
        //			sQry = sQry + "'" + Comments + "'";
        //			//비고

        //			RecordSet02.DoQuery(sQry);

        //			MDC_Com.MDC_GF_Message(ref "등록 완료!", ref "S");

        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet02 = null;
        //			functionReturnValue = true;
        //			return functionReturnValue;
        //			PS_MM004_AddData_Error:

        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			functionReturnValue = false;
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet02 = null;
        //			MDC_Com.MDC_GF_Message(ref "PS_MM004_AddData_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //			return functionReturnValue;
        //		}
        #endregion

        #region PS_MM004_CheckDataValie
        //		private bool PS_MM004_CheckDataValie()
        //		{
        //			bool functionReturnValue = false;
        //			//******************************************************************************
        //			//Function ID : PS_MM004_CheckDataValie()
        //			//해당모듈 : PS_MM004
        //			//기능 : 필수입력사항 체크
        //			//인수 : 없음
        //			//반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
        //			//특이사항 : 없음
        //			//******************************************************************************
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			short ErrNum = 0;
        //			ErrNum = 0;

        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			switch (true) {
        //				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("BPLId").Specific.Value)):
        //					//사업장
        //					ErrNum = 1;
        //					goto PS_MM004_CheckDataValie_Error;
        //					break;
        //				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ItmBsort").Specific.Value)):
        //					//품목분류
        //					ErrNum = 2;
        //					goto PS_MM004_CheckDataValie_Error;
        //					break;
        //				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("LocCode").Specific.Value)):
        //					//운송지역코드
        //					ErrNum = 3;
        //					goto PS_MM004_CheckDataValie_Error;
        //					break;
        //				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("Way").Specific.Value)):
        //					//편도/회로
        //					ErrNum = 4;
        //					goto PS_MM004_CheckDataValie_Error;
        //					break;
        //				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("TranCard").Specific.Value)):
        //					//운송업체
        //					ErrNum = 5;
        //					goto PS_MM004_CheckDataValie_Error;
        //					break;
        //				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("TranCode").Specific.Value)):
        //					//차량번호
        //					ErrNum = 6;
        //					goto PS_MM004_CheckDataValie_Error;
        //					break;
        //				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("Tonnage").Specific.Value)):
        //					//차종(톤수)
        //					ErrNum = 7;
        //					goto PS_MM004_CheckDataValie_Error;
        //					break;
        //				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ItemName").Specific.Value)):
        //					//품명
        //					ErrNum = 8;
        //					goto PS_MM004_CheckDataValie_Error;
        //					break;
        //			}

        //			functionReturnValue = true;
        //			return functionReturnValue;
        //			PS_MM004_CheckDataValie_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			if (ErrNum == 1) {
        //				MDC_Com.MDC_GF_Message(ref "사업장은 필수사항입니다. 확인하세요.", ref "E");
        //				oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else if (ErrNum == 2) {
        //				MDC_Com.MDC_GF_Message(ref "품목분류는 필수사항입니다. 확인하세요.", ref "E");
        //				oForm.Items.Item("ItmBsort").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else if (ErrNum == 3) {
        //				MDC_Com.MDC_GF_Message(ref "운송지역코드는 필수사항입니다. 확인하세요.", ref "E");
        //				oForm.Items.Item("LocCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else if (ErrNum == 4) {
        //				MDC_Com.MDC_GF_Message(ref "편도/회로는 필수사항입니다. 확인하세요.", ref "E");
        //				oForm.Items.Item("Way").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else if (ErrNum == 5) {
        //				MDC_Com.MDC_GF_Message(ref "운송업체는 필수사항입니다. 확인하세요.", ref "E");
        //				oForm.Items.Item("TranCard").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else if (ErrNum == 6) {
        //				MDC_Com.MDC_GF_Message(ref "차량번호는 필수사항입니다. 확인하세요.", ref "E");
        //				oForm.Items.Item("TranCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else if (ErrNum == 7) {
        //				MDC_Com.MDC_GF_Message(ref "차종(톤수)은 필수사항입니다. 확인하세요.", ref "E");
        //				oForm.Items.Item("Tonnage").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else if (ErrNum == 8) {
        //				MDC_Com.MDC_GF_Message(ref "품명은 필수사항입니다. 확인하세요.", ref "E");
        //				oForm.Items.Item("Tonnage").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			}
        //			functionReturnValue = false;
        //			return functionReturnValue;
        //		}
        #endregion



        #region PS_MM004_Print_Report01
        //		private void PS_MM004_Print_Report01()
        //		{

        //			string DocNum = null;
        //			short ErrNum = 0;
        //			string WinTitle = null;
        //			string ReportName = null;
        //			string sQry = null;
        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			int DocEntry = 0;

        //			oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			/// ODBC 연결 체크
        //			MDC_PS_Common.ConnectODBC();

        //			////인자 MOVE , Trim 시키기..
        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			DocEntry = Convert.ToInt32(Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value));

        //			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

        //			WinTitle = "[PS_MM004] 운송확인증";

        //			ReportName = "PS_MM004_10.rpt";
        //			MDC_Globals.gRpt_Formula = new string[3];
        //			MDC_Globals.gRpt_Formula_Value = new string[3];
        //			MDC_Globals.gRpt_SRptSqry = new string[2];
        //			MDC_Globals.gRpt_SRptName = new string[2];
        //			MDC_Globals.gRpt_SFormula = new string[2, 2];
        //			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];


        //			//// Formula 수식필드

        //			//// SubReport


        //			MDC_Globals.gRpt_SFormula[1, 1] = "";
        //			MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

        //			/// Procedure 실행"
        //			sQry = "EXEC [PS_MM004_10] '" + DocEntry + "'";

        //			oRecordSet.DoQuery(sQry);
        //			if (oRecordSet.RecordCount == 0) {
        //				ErrNum = 1;
        //				goto PS_MM004_Print_Report01_Error;
        //			}

        //			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") == false) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			}

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return;
        //			PS_MM004_Print_Report01_Error:

        //			if (ErrNum == 1) {
        //				//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oRecordSet = null;
        //				MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다. 확인해 주세요.", ref "E");
        //			} else {
        //				//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oRecordSet = null;
        //				SubMain.Sbo_Application.SetStatusBarMessage("PS_MM004_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			}

        //		}
        #endregion











        #region Raise_ItemEvent
        /////아이템 변경 이벤트
        //		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement



        //			switch (pVal.EventType) {
        //				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //					////1
        //					Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //					////2
        //					if (pVal.CharPressed == 9) {
        //						if (pVal.ItemUID == "ItmBsort") {
        //							//UPGRADE_WARNING: oForm.Items(ItmBsort).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (string.IsNullOrEmpty(oForm.Items.Item("ItmBsort").Specific.Value)) {
        //								SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //								BubbleEvent = false;
        //							}
        //						///운송지역
        //						} else if (pVal.ItemUID == "LocCode") {
        //							//UPGRADE_WARNING: oForm.Items(LocCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (string.IsNullOrEmpty(oForm.Items.Item("LocCode").Specific.Value)) {
        //								SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //								BubbleEvent = false;
        //							}

        //						//차종(톤수)
        //						} else if (pVal.ItemUID == "Tonnage") {
        //							//UPGRADE_WARNING: oForm.Items(Tonnage).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (string.IsNullOrEmpty(oForm.Items.Item("Tonnage").Specific.Value)) {
        //								SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //								BubbleEvent = false;
        //							}
        //						}
        //					}

        //					Raise_EVENT_KEY_DOWN(ref FormUID, ref pVal, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //					////5
        //					Raise_EVENT_COMBO_SELECT(ref FormUID, ref pVal, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_CLICK:
        //					////6
        //					Raise_EVENT_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //					////7
        //					Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //					////8
        //					Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //					////10
        //					if (pVal.ItemChanged == true) {
        //						if (pVal.ItemUID == "ItmBsort") {
        //							PS_MM004_FlushToItemValue(pVal.ItemUID);
        //						} else if (pVal.ItemUID == "LocCode") {
        //							PS_MM004_FlushToItemValue(pVal.ItemUID);
        //						} else if (pVal.ItemUID == "Tonnage") {
        //							PS_MM004_FlushToItemValue(pVal.ItemUID);
        //						} else if (pVal.ItemUID == "Amt") {
        //							PS_MM004_FlushToItemValue(pVal.ItemUID);
        //						} else if (pVal.ItemUID == "PassAmt") {
        //							PS_MM004_FlushToItemValue(pVal.ItemUID);
        //						} else if (pVal.ItemUID == "TWeight") {
        //							PS_MM004_FlushToItemValue(pVal.ItemUID);
        //						} else if (pVal.ItemUID == "DocDate") {
        //							PS_MM004_FlushToItemValue(pVal.ItemUID);
        //						}
        //					}
        //					break;
        //				// Call Raise_EVENT_VALIDATE(FormUID, pVal, BubbleEvent)
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //					////11
        //					Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pVal, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //					////18
        //					break;
        //				////et_FORM_ACTIVATE
        //				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //					////19
        //					break;
        //				////et_FORM_DEACTIVATE
        //				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //					////20
        //					Raise_EVENT_RESIZE(ref FormUID, ref pVal, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //					////27
        //					Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pVal, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //					////3
        //					Raise_EVENT_GOT_FOCUS(ref FormUID, ref pVal, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //					////4
        //					break;
        //				////et_LOST_FOCUS
        //				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //					////17
        //					Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pVal, ref BubbleEvent);
        //					break;
        //			}
        //			return;
        //			Raise_ItemEvent_Error:
        //			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_MenuEvent
        //		public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////BeforeAction = True
        //			if ((pVal.BeforeAction == true)) {
        //				switch (pVal.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1293":
        //						//행삭제
        //						break;

        //					case "1281":
        //						//찾기
        //						break;
        //					case "1282":
        //						//추가
        //						///추가버튼 클릭시 메트릭스 insertrow
        //						PS_MM004_SetForm();

        //						//                oMat01.Clear
        //						//                oMat01.FlushToDataSource
        //						//                oMat01.LoadFromDataSource

        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        //						BubbleEvent = false;
        //						PS_MM004_LoadCaption();
        //						return;

        //						break;
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;
        //				}
        //			////BeforeAction = False
        //			} else if ((pVal.BeforeAction == false)) {
        //				switch (pVal.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1293":
        //						//행삭제
        //						break;

        //					case "1281":
        //						//찾기
        //						break;
        //					////Call PS_MM004_PS_MM004_FormItemEnabled '//UDO방식
        //					case "1282":
        //						//추가
        //						break;
        //					//                oMat01.Clear
        //					//                oDS_PS_MM004H.Clear

        //					//                Call PS_MM004_LoadCaption
        //					//                Call PS_MM004_FormItemEnabled
        //					////Call PS_MM004_PS_MM004_FormItemEnabled '//UDO방식
        //					////Call PS_MM004_AddMatrixRow(0, True) '//UDO방식
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;
        //					////Call PS_MM004_PS_MM004_FormItemEnabled
        //				}
        //			}
        //			return;
        //			Raise_MenuEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormDataEvent
        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////BeforeAction = True
        //			if ((BusinessObjectInfo.BeforeAction == true)) {
        //				switch (BusinessObjectInfo.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //						////33
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //						////34
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //						////36
        //						break;
        //				}
        //			////BeforeAction = False
        //			} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //				switch (BusinessObjectInfo.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //						////33
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //						////34
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //						////36
        //						break;
        //				}
        //			}
        //			return;
        //			Raise_FormDataEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_RightClickEvent
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pVal.BeforeAction == true) {
        //			} else if (pVal.BeforeAction == false) {
        //			}
        //			if (pVal.ItemUID == "Mat01") {
        //				if (pVal.Row > 0) {
        //					oLastItemUID01 = pVal.ItemUID;
        //					oLastColUID01 = pVal.ColUID;
        //					oLastColRow01 = pVal.Row;
        //				}
        //			} else {
        //				oLastItemUID01 = pVal.ItemUID;
        //				oLastColUID01 = "";
        //				oLastColRow01 = 0;
        //			}
        //			return;
        //			Raise_RightClickEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pVal.BeforeAction == true) {
        //				if (pVal.ItemUID == "PS_MM004") {
        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					}
        //				}

        //				///추가/확인 버튼클릭
        //				if (pVal.ItemUID == "BtnAdd") {
        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //						if (PS_MM004_CheckDataValie() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}
        //						if (PS_MM004_AddData() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}

        //						PS_MM004_SetForm();
        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

        //						PS_MM004_LoadCaption();
        //						PS_MM004_MTX01();

        //						oLast_Mode = oForm.Mode;

        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //						if (PS_MM004_CheckDataValie() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}
        //						if (PS_MM004_UpdateData() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}

        //						PS_MM004_SetForm();
        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

        //						PS_MM004_LoadCaption();
        //						PS_MM004_MTX01();
        //					}

        //				///조회
        //				} else if (pVal.ItemUID == "BtnSearch") {
        //					//            Call PS_MM004_SetForm
        //					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        //					///fm_VIEW_MODE

        //					PS_MM004_LoadCaption();
        //					PS_MM004_MTX01();

        //				///삭제
        //				} else if (pVal.ItemUID == "BtnDelete") {
        //					if (SubMain.Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1")) {

        //						PS_MM004_DeleteData();
        //						PS_MM004_SetForm();
        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        //						///fm_VIEW_MODE

        //						PS_MM004_LoadCaption();
        //						PS_MM004_MTX01();

        //					} else {

        //					}
        //				} else if (pVal.ItemUID == "BtnPrint") {

        //					PS_MM004_Print_Report01();


        //				}


        //			} else if (pVal.BeforeAction == false) {
        //				if (pVal.ItemUID == "PS_MM004") {
        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					}
        //				}
        //			}
        //			return;
        //			Raise_EVENT_ITEM_PRESSED_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pVal.BeforeAction == true) {
        //				//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "ItemCode", "") '//사용자값활성
        //				//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "Mat01", "FixCode") '자산번호 포맷서치
        //			} else if (pVal.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_KEY_DOWN_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_CLICK
        //		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			string DocEntry = null;
        //			string sQry = null;
        //			SAPbobsCOM.Recordset oRecordSet01 = null;

        //			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			if (pVal.BeforeAction == true) {
        //				if (pVal.ItemUID == "Mat01") {
        //					if (pVal.Row > 0) {

        //						oMat01.SelectRow(pVal.Row, true, false);

        //						//                Call oForm.Freeze(True)

        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						DocEntry = oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value;

        //						PS_MM004_RET(DocEntry);

        //						//                sQry = "Select DocEntry, U_ObjType, U_DocNo, U_BPLId, U_DocDate, U_ItmBsort, U_ItmBName, U_LocCode, U_LocName, "
        //						//                sQry = sQry + " U_Way, U_SectionF, U_SectionT, U_Destin, U_TranCard, U_TranCode, U_Tonnage, U_Weight, U_Amt, U_Comments "
        //						//                sQry = sQry + " From [@PS_MM004H] Where DocEntry = '" + DocEntry + "'"
        //						//
        //						//                oRecordSet01.DoQuery sQry
        //						//
        //						//                'DataSource를 이용하여 각 컨트롤에 값을 출력
        //						//                Call oDS_PS_MM004H.setValue("DocEntry", 0, oRecordSet01.Fields("DocEntry").Value) '관리번호
        //						//                Call oDS_PS_MM004H.setValue("U_ObjType", 0, Trim(oRecordSet01.Fields("U_ObjType").Value)) '참조Object
        //						//                Call oDS_PS_MM004H.setValue("U_DocNo", 0, oRecordSet01.Fields("U_DocNo").Value) '참조문서번호
        //						//                Call oDS_PS_MM004H.setValue("U_BPLId", 0, Trim(oRecordSet01.Fields("U_BPLId").Value)) '사업장
        //						//                Call oDS_PS_MM004H.setValue("U_DocDate", 0, Format(oRecordSet01.Fields("U_DocDate").Value, "YYYYMMDD")) '일자
        //						//                Call oDS_PS_MM004H.setValue("U_ItmBsort", 0, Trim(oRecordSet01.Fields("U_ItmBsort").Value)) '분류코드
        //						//                Call oDS_PS_MM004H.setValue("U_ItmBName", 0, Trim(oRecordSet01.Fields("U_ItmBName").Value)) '분류명
        //						//                Call oDS_PS_MM004H.setValue("U_LocCode", 0, Trim(oRecordSet01.Fields("U_LocCode").Value)) '운송지역코드
        //						//                Call oDS_PS_MM004H.setValue("U_LocName", 0, Trim(oRecordSet01.Fields("U_LocName").Value)) '운송지역명
        //						//                Call oDS_PS_MM004H.setValue("U_Way", 0, Trim(oRecordSet01.Fields("U_Way").Value)) '편도/회로
        //						//                Call oDS_PS_MM004H.setValue("U_SectionF", 0, Trim(oRecordSet01.Fields("U_SectionF").Value)) '운행구간시작
        //						//                Call oDS_PS_MM004H.setValue("U_SectionT", 0, Trim(oRecordSet01.Fields("U_SectionT").Value)) '운행구간종료
        //						//                Call oDS_PS_MM004H.setValue("U_Destin", 0, Trim(oRecordSet01.Fields("U_Destin").Value)) '도착장소
        //						//                Call oDS_PS_MM004H.setValue("U_TranCard", 0, Trim(oRecordSet01.Fields("U_TranCard").Value)) '운송업체
        //						//                Call oDS_PS_MM004H.setValue("U_TranCode", 0, Trim(oRecordSet01.Fields("U_TranCode").Value)) '차량번호
        //						//                Call oDS_PS_MM004H.setValue("U_Tonnage", 0, oRecordSet01.Fields("U_Tonnage").Value) '차종(톤수)
        //						//                Call oDS_PS_MM004H.setValue("U_Weight", 0, oRecordSet01.Fields("U_Weight").Value) '운송량
        //						//                Call oDS_PS_MM004H.setValue("U_Amt", 0, oRecordSet01.Fields("U_Amt").Value) '금액
        //						//                Call oDS_PS_MM004H.setValue("U_Comments", 0, Trim(oRecordSet01.Fields("U_Comments").Value)) '비고
        //						//
        //						//                oForm.Mode = fm_UPDATE_MODE
        //						//                Call PS_MM004_LoadCaption
        //						//
        //						//                Call oForm.Freeze(False)

        //					}
        //				}
        //			} else if (pVal.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_CLICK_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //		private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			int i = 0;
        //			string sQry = null;

        //			int sCount = 0;
        //			int sSeq = 0;
        //			string sCode = null;
        //			string SCpCode = null;
        //			double TWeight = 0;

        //			SAPbobsCOM.Recordset oRecordSet01 = null;

        //			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			oForm.Freeze(true);
        //			if (pVal.BeforeAction == true) {

        //			} else if (pVal.BeforeAction == false) {

        //				if (pVal.ItemUID == "Way") {
        //					//UPGRADE_WARNING: oForm.Items(TWeight).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (!string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("LocCode").Specific.Value)) & oForm.Items.Item("TWeight").Specific.Value > 0) {
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						TWeight = oForm.Items.Item("TWeight").Specific.Value;

        //						sQry = "Select Amt = Isnull(b.U_Amt,0) From [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						sQry = sQry + " Where a.U_BPLId =  '" + Strings.Trim(oForm.Items.Item("BPLId").Specific.Value) + "'";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						sQry = sQry + "   And b.U_LocCode = '" + Strings.Trim(oForm.Items.Item("LocCode").Specific.Value) + "'";
        //						sQry = sQry + "   And " + TWeight + " Between b.U_KgF And b.U_KgT ";
        //						oRecordSet01.DoQuery(sQry);

        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (Strings.Trim(oForm.Items.Item("Way").Specific.Value) == "10") {
        //							//UPGRADE_WARNING: oForm.Items(Amt).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oForm.Items.Item("Amt").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
        //							//                        oForm.Items("TotAmt").Specific.Value = Trim(oRecordSet01.Fields(0).Value) + oForm.Items("PassAmt").Specific.Value
        //							//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						} else if (Strings.Trim(oForm.Items.Item("Way").Specific.Value) == "20") {
        //							//UPGRADE_WARNING: oForm.Items(Amt).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oForm.Items.Item("Amt").Specific.Value = (oRecordSet01.Fields.Item(0).Value * 0.8);
        //							//                        oForm.Items("TotAmt").Specific.Value = (oRecordSet01.Fields(0).Value * 0.8) + oForm.Items("PassAmt").Specific.Value
        //						}

        //					}

        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (Strings.Trim(oForm.Items.Item("Way").Specific.Value) == "10") {
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("SectionF").Specific.Value = "";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("SectionT").Specific.Value = "";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("ItemName").Specific.Value = "";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("Destin").Specific.Value = "";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (Strings.Trim(oForm.Items.Item("Way").Specific.Value) == "20") {
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("ItmBsort").Specific.Value = "102";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("LocCode").Specific.Value = "07";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("SectionF").Specific.Value = "안강";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("SectionT").Specific.Value = "창원";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("ItemName").Specific.Value = "부품소재";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("Destin").Specific.Value = "창원";
        //					}

        //				} else if (pVal.ItemUID == "PassYN") {
        //					//UPGRADE_WARNING: oForm.Items(PassYN).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oForm.Items.Item("PassYN").Specific.Value == "Y") {
        //						//경유지(혼적) Y이면
        //						sQry = "Select PassAmt = Isnull(U_PassAmt,0) From [@PS_MM003H] a Inner Join [@PS_MM003L] b On a.Code = b.Code";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						sQry = sQry + " Where a.U_BPLId =  '" + Strings.Trim(oForm.Items.Item("BPLId").Specific.Value) + "'";
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						sQry = sQry + "   And b.U_LocCode = '" + Strings.Trim(oForm.Items.Item("LocCode").Specific.Value) + "'";
        //						oRecordSet01.DoQuery(sQry);

        //						//UPGRADE_WARNING: oForm.Items(PassAmt).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oRecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("PassAmt").Specific.Value = oRecordSet01.Fields.Item(0).Value;
        //						//                    oForm.Items("TotAmt").Specific.Value = oForm.Items("Amt").Specific.Value + oRecordSet01.Fields(0).Value

        //					} else {
        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("PassAmt").Specific.Value = 0;
        //						//                    oForm.Items("TotAmt").Specific.Value = oForm.Items("Amt").Specific.Value
        //					}
        //				}

        //				//
        //				//
        //				//                sCount = oForm.Items("CpCode").Specific.ValidValues.Count
        //				//                sSeq = sCount
        //				//                For i = 1 To sCount
        //				//                    oForm.Items("CpCode").Specific.ValidValues.Remove sSeq - 1, psk_Index
        //				//                    sSeq = sSeq - 1
        //				//                Next i
        //				//
        //				//                '//공정구분에 따른 공정코드변경
        //				//
        //				//                Select Case oForm.Items("CpGbn").Specific.Value
        //				//                    Case "10"
        //				//                        '//멀티게이지-금형래핑
        //				//                        Call oForm.Items("WorkGbn").Specific.Select("104", psk_ByValue)
        //				//                        oForm.Items("ItmBsort").Enabled = True
        //				//                        oMat01.Columns("ItmBsort").Editable = True
        //				//                    Case "20"
        //				//                        '//멀티게이지-외경연삭
        //				//                        Call oForm.Items("WorkGbn").Specific.Select("104", psk_ByValue)
        //				//                        oForm.Items("ItmBsort").Enabled = False
        //				//                        oMat01.Columns("ItmBsort").Editable = False
        //				//                    Case "30"
        //				//                        '//멀티게이지-포장
        //				//                        Call oForm.Items("WorkGbn").Specific.Select("104", psk_ByValue)
        //				//                        oForm.Items("ItmBsort").Enabled = False
        //				//                        oMat01.Columns("ItmBsort").Editable = False
        //				//                    Case "40"
        //				//                        '//휘팅-바렐
        //				//                        Call oForm.Items("WorkGbn").Specific.Select("101", psk_ByValue)
        //				//                        oForm.Items("ItmBsort").Enabled = False
        //				//                        oMat01.Columns("ItmBsort").Editable = False
        //				//                    Case "50"
        //				//                        '//휘팅-포장
        //				//                        Call oForm.Items("WorkGbn").Specific.Select("101", psk_ByValue)
        //				//                        oForm.Items("ItmBsort").Enabled = False
        //				//                        oMat01.Columns("ItmBsort").Editable = False
        //				//                    Case "60", "70"
        //				//                        '//검사공수 입력
        //				//                        Call oForm.Items("WorkGbn").Specific.Select("105", psk_ByValue)
        //				//                    Case Else
        //				//
        //				//                End Select
        //				//
        //				//            ElseIf pVal.ItemUID = "OrdType" Then
        //				//
        //				//                If oForm.Items("OrdType").Specific.Value = "10" Then
        //				//                    '// 실동입력
        //				//                    oForm.Items("CpGbn").Enabled = True
        //				//'                    oForm.Items("CntcCode").Enabled = True '사번 TextBox 활성, 비활성 기능 제외(2013.05.03 송명규 추가)
        //				//                    oForm.Items("CpCode").Enabled = True
        //				//                    oForm.Items("ItmBsort").Enabled = True
        //				//                    oMat01.Columns("NCode").Editable = False
        //				//                Else
        //				//                    '// 비가동입력
        //				//                    oForm.Items("CpGbn").Enabled = False
        //				//
        //				//                    Call oForm.Items("CpGbn").Specific.Select("", psk_ByValue)
        //				//'                    oForm.Items("CntcCode").Enabled = False '사번 TextBox 활성, 비활성 기능 제외(2013.05.03 송명규 추가)
        //				//
        //				//                    oForm.Items("CpCode").Enabled = False
        //				//                    Call oForm.Items("CpCode").Specific.Select("", psk_ByValue)
        //				//                    oForm.Items("ItmBsort").Enabled = False
        //				//                    oMat01.Columns("NCode").Editable = True
        //				//
        //				//
        //				//                End If
        //				//            ElseIf pVal.ItemUID = "CpCode" Then '// 공정코드
        //				//                sQry = "SELECT U_CpName From [@PS_PP001L] Where U_CpCode = '" & Trim(oForm.Items("CpCode").Specific.Value) & "'"
        //				//                oRecordSet01.DoQuery sQry
        //				//                oForm.Items("CpName").Specific.Value = Trim(oRecordSet01.Fields(0).Value)
        //				//            End If



        //			}
        //			oForm.Freeze(false);
        //			return;
        //			Raise_EVENT_COMBO_SELECT_Error:
        //			oForm.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
        //		private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pVal.BeforeAction == true) {

        //			} else if (pVal.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_DOUBLE_CLICK_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
        //		private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pVal.BeforeAction == true) {

        //			} else if (pVal.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion


        #region Raise_EVENT_VALIDATE
        //		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			oForm.Freeze(true);
        //			if (pVal.BeforeAction == true) {
        //				//        If pVal.ItemChanged = True Then
        //				//            If (pVal.ItemUID = "Mat01") Then
        //				//                If (pVal.ColUID = "ItemCode") Then
        //				//                    '//기타작업
        //				//                    Call oDS_PS_MM004L.setValue("U_" & pVal.ColUID, pVal.Row - 1, oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value)
        //				//                    If oMat01.RowCount = pVal.Row And Trim(oDS_PS_MM004L.GetValue("U_" & pVal.ColUID, pVal.Row - 1)) <> "" Then
        //				//                        PS_MM004_AddMatrixRow (pVal.Row)
        //				//                    End If
        //				//                Else
        //				//                    Call oDS_PS_MM004L.setValue("U_" & pVal.ColUID, pVal.Row - 1, oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value)
        //				//                End If
        //				//            Else
        //				//                If (pVal.ItemUID = "DocEntry") Then
        //				//                    Call oDS_PS_MM004H.setValue(pVal.ItemUID, 0, oForm.Items(pVal.ItemUID).Specific.Value)
        //				//                ElseIf (pVal.ItemUID = "CardCode") Then
        //				//                    Call oDS_PS_MM004H.setValue("U_" & pVal.ItemUID, 0, oForm.Items(pVal.ItemUID).Specific.Value)
        //				//                    Call oDS_PS_MM004H.setValue("U_CardName", 0, MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" & oForm.Items(pVal.ItemUID).Specific.Value & "'"))
        //				//                Else
        //				//                    Call oDS_PS_MM004H.setValue("U_" & pVal.ItemUID, 0, oForm.Items(pVal.ItemUID).Specific.Value)
        //				//                End If
        //				//            End If
        //				//            oMat01.LoadFromDataSource
        //				//            oMat01.AutoResizeColumns
        //				//            oForm.Update
        //				//        End If
        //			} else if (pVal.BeforeAction == false) {

        //			}
        //			oForm.Freeze(false);
        //			return;
        //			Raise_EVENT_VALIDATE_Error:
        //			oForm.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //		private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pVal.BeforeAction == true) {

        //			} else if (pVal.BeforeAction == false) {
        //				PS_MM004_PS_MM004_FormItemEnabled();
        //				////Call PS_MM004_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LOAD_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_RESIZE
        //		private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pVal = null, ref bool BubbleEvent = false)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pVal.BeforeAction == true) {

        //			} else if (pVal.BeforeAction == false) {
        //				PS_MM004_FormResize();
        //			}
        //			return;
        //			Raise_EVENT_RESIZE_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //		private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pVal.BeforeAction == true) {

        //			} else if (pVal.BeforeAction == false) {
        //				//        If (pVal.ItemUID = "ItemCode") Then
        //				//            Dim oDataTable01 As SAPbouiCOM.DataTable
        //				//            Set oDataTable01 = pVal.SelectedObjects
        //				//            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
        //				//            Set oDataTable01 = Nothing
        //				//        End If
        //				//        If (pVal.ItemUID = "CardCode" Or pVal.ItemUID = "CardName") Then
        //				//            Call MDC_GP_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_MM004H", "U_CardCode,U_CardName")
        //				//        End If
        //			}
        //			return;
        //			Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pVal.ItemUID == "Mat01") {
        //				if (pVal.Row > 0) {
        //					oLastItemUID01 = pVal.ItemUID;
        //					oLastColUID01 = pVal.ColUID;
        //					oLastColRow01 = pVal.Row;
        //				}
        //			} else {
        //				oLastItemUID01 = pVal.ItemUID;
        //				oLastColUID01 = "";
        //				oLastColRow01 = 0;
        //			}
        //			return;
        //			Raise_EVENT_GOT_FOCUS_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pVal.BeforeAction == true) {
        //			} else if (pVal.BeforeAction == false) {
        //				SubMain.RemoveForms(oFormUniqueID);
        //				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oForm = null;
        //				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oMat01 = null;
        //			}
        //			return;
        //			Raise_EVENT_FORM_UNLOAD_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion




    }
}

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
        private void LoadForm()
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
                PS_MM004_SetForm("");
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
                    oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("SDocDateF").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                }

                oForm.Items.Item("SDocDateT").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
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
        private void PS_MM004_SetForm(string mode)
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

                //==========조회정보==========
                if (mode != "U") //업데이트모드일때는 조회정보 초기화 안함(수정한 자료 조회)
                {
                    oForm.Items.Item("SDocDateF").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("SDocDateT").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("SLocCode").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue); //운송지역
                    oForm.Items.Item("SWay").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue); //편도/회로
                    oForm.Items.Item("STonnage").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue); //차종(톤수)
                    oForm.Items.Item("SObjType").Specific.Value = "";
                    oForm.Items.Item("SDocNo").Specific.Value = "";
                }

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

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PS_MM004_MTX01()
        {
            short i;
            string sQry;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string SBPLID; //사업장
            string SDocDateF; //조회기간1
            string SDocDateT; //조회기간2
            string SItmBsort; //폼목분류
            string SLocCode; //운송지역
            string SWay; //편도/회로
            string STranCard; //운송업체
            string STonnage; //차종(톤수)
            string SObjType; //참조Object
            string SDocNo; //참조문서번호

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
                oForm.Freeze(true);

                SBPLID = oForm.Items.Item("SBPLId").Specific.Value.ToString().Trim(); //사업장
                SDocDateF = oForm.Items.Item("SDocDateF").Specific.Value.ToString().Trim(); //조회시작일
                SDocDateT = oForm.Items.Item("SDocDateT").Specific.Value.ToString().Trim(); //조회종료일
                SItmBsort = oForm.Items.Item("SItmBsort").Specific.Value.ToString().Trim(); //품목분류
                SLocCode = oForm.Items.Item("SLocCode").Specific.Value.ToString().Trim(); //운송지역
                SWay = oForm.Items.Item("SWay").Specific.Value.ToString().Trim(); //편도/회로
                STranCard = oForm.Items.Item("STranCard").Specific.Value.ToString().Trim(); //운송업체
                STonnage = oForm.Items.Item("STonnage").Specific.Value.ToString().Trim(); //차종(톤수)
                SObjType = oForm.Items.Item("SObjType").Specific.Value.ToString().Trim(); //참조Object
                SDocNo = oForm.Items.Item("SDocNo").Specific.Value.ToString().Trim(); //참조문서번호

                sQry = " EXEC [PS_MM004_01] ";
                sQry += "'" + SBPLID + "',"; //사업장
                sQry += "'" + SDocDateF + "',"; //조회시작일
                sQry += "'" + SDocDateT + "',"; //조회종료일
                sQry += "'" + SItmBsort + "',"; //품목분류
                sQry += "'" + SLocCode + "',"; //운송지역
                sQry += "'" + SWay + "',"; //편도/회로
                sQry += "'" + STranCard + "',"; //운송업체
                sQry += "'" + STonnage + "',"; //차종(톤수)
                sQry += "'" + SObjType + "',"; //참조Object
                sQry += "'" + SDocNo + "'"; //참조문서번호

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_MM004L.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다.확인하세요.";
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PS_MM004_LoadCaption();
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_MM004L.Size)
                    {
                        oDS_PS_MM004L.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_MM004L.Offset = i;

                    oDS_PS_MM004L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_MM004L.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim()); //관리번호
                    oDS_PS_MM004L.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("BPLId").Value.ToString().Trim()); //사업장
                    oDS_PS_MM004L.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("DocDate").Value.ToString().Trim()); //일자
                    oDS_PS_MM004L.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("ItmBName").Value.ToString().Trim()); //품목분류명
                    oDS_PS_MM004L.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("LocName").Value.ToString().Trim()); //운송지역
                    oDS_PS_MM004L.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("Way").Value.ToString().Trim()); //편도/회로
                    oDS_PS_MM004L.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("SectionF").Value.ToString().Trim()); //운행구간(출발)
                    oDS_PS_MM004L.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("SectionT").Value.ToString().Trim()); //운행구간(도착)
                    oDS_PS_MM004L.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("Destin").Value.ToString().Trim()); //도착장소
                    oDS_PS_MM004L.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("TranCard").Value.ToString().Trim()); //운송업체
                    oDS_PS_MM004L.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("TranCode").Value.ToString().Trim()); //차량번호
                    oDS_PS_MM004L.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("Tonnage").Value.ToString().Trim()); //차종(톤수)
                    oDS_PS_MM004L.SetValue("U_ColReg13", i, oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim()); //품명
                    oDS_PS_MM004L.SetValue("U_ColNum01", i, oRecordSet01.Fields.Item("Weight").Value.ToString().Trim()); //운송량
                    oDS_PS_MM004L.SetValue("U_ColSum01", i, oRecordSet01.Fields.Item("TotAmt").Value.ToString().Trim()); //금액
                    oDS_PS_MM004L.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("PassYN").Value.ToString().Trim()); 
                    oDS_PS_MM004L.SetValue("U_ColReg16", i, oRecordSet01.Fields.Item("Comments").Value.ToString().Trim()); //비고

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
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 기본정보 삭제
        /// </summary>
        private void PS_MM004_DeleteData()
        {
            string sQry;
            string errMessage = string.Empty;
            string DocEntry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                    sQry = "SELECT COUNT(*) FROM [@PS_MM004H] WHERE DocEntry = '" + DocEntry + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (oRecordSet01.RecordCount == 0)
                    {
                        errMessage = "삭제대상이 없습니다. 확인하세요.";
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        throw new Exception();
                    }
                    else
                    {
                        sQry = "EXEC PS_MM004_04 '" + DocEntry + "'";
                        oRecordSet01.DoQuery(sQry);
                    }
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 기본정보 수정
        /// </summary>
        /// <returns></returns>
        private bool PS_MM004_UpdateData()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            string sQry;
            int DocEntry; //관리번호
            string BPLId; //사업장
            string DocDate; //일자
            string ItmBsort; //분류코드
            string ItmBName; //분류명
            string LocCode; //운송지역코드
            string LocName; //운송지역명
            string Way; //편도/회로
            string SectionF; //운행구간출발
            string SectionT; //운행구간도착
            string Destin; //도착장소
            string TranCard; //운송업체
            string TranCode; //차량번호
            double Tonnage; //차종(톤수)
            string ItemName; //품명
            double Weight; //실중량
            double TWeight; //총중량
            string PassYN; //경유지YN
            double Amt; //기본운임
            double PassAmt; //경유지운임
            double TotAmt; //총운임
            string Comments; //비고
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEntry = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim()); //관리번호
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(); //사업장
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim(); //일자
                ItmBsort = oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim(); //품목분류코드
                ItmBName = oForm.Items.Item("ItmBName").Specific.Value.ToString().Trim(); //품목분류명
                LocCode = oForm.Items.Item("LocCode").Specific.Value.ToString().Trim(); //운송지역코드
                LocName = oForm.Items.Item("LocName").Specific.Value.ToString().Trim(); //운송지역명
                Way = oForm.Items.Item("Way").Specific.Value.ToString().Trim(); //편도/회로
                SectionF = oForm.Items.Item("SectionF").Specific.Value.ToString().Trim(); //운행구간출발
                SectionT = oForm.Items.Item("SectionT").Specific.Value.ToString().Trim(); //운행구간도착
                Destin = oForm.Items.Item("Destin").Specific.Value.ToString().Trim(); //도착장소
                TranCard = oForm.Items.Item("TranCard").Specific.Value.ToString().Trim(); //운송업체
                TranCode = oForm.Items.Item("TranCode").Specific.Value.ToString().Trim(); //차량번호
                Tonnage = Convert.ToDouble(oForm.Items.Item("Tonnage").Specific.Value); //차종(톤수)
                ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim(); //품명
                Weight = Convert.ToDouble(oForm.Items.Item("Weight").Specific.Value); //실중량
                TWeight = Convert.ToDouble(oForm.Items.Item("TWeight").Specific.Value); //계근중량
                PassYN = oForm.Items.Item("PassYN").Specific.Value; //경유지YN
                Amt = Convert.ToDouble(oForm.Items.Item("Amt").Specific.Value); //기본운임
                PassAmt = Convert.ToDouble(oForm.Items.Item("PassAmt").Specific.Value); //경유지운임
                TotAmt = Convert.ToDouble(oForm.Items.Item("TotAmt").Specific.Value); //총운임
                Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim(); //비고

                if (string.IsNullOrEmpty(Convert.ToString(DocEntry).ToString().Trim()))
                {
                    errMessage = "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!";
                    throw new Exception();
                }

                sQry = " EXEC [PS_MM004_03] ";
                sQry += DocEntry + ","; //관리번호
                sQry += "'" + BPLId + "',"; //사업장
                sQry += "'" + DocDate + "',"; //출장번호1
                sQry += "'" + ItmBsort + "',"; //출장번호2
                sQry += "'" + ItmBName + "',"; //사원번호
                sQry += "'" + LocCode + "',"; //사원성명
                sQry += "'" + LocName + "',"; //출장지
                sQry += "'" + Way + "',"; //출장지상세
                sQry += "'" + SectionF + "',"; //작번
                sQry += "'" + SectionT + "',"; //시작일자
                sQry += "'" + Destin + "',"; //시작시각
                sQry += "'" + TranCard + "',"; //종료일자
                sQry += "'" + TranCode + "',"; //종료시각
                sQry += Tonnage + ","; //차종(톤수)
                sQry += "'" + ItemName + "',"; //품명
                sQry += Weight + ","; //실중량
                sQry += TWeight + ","; //총중량
                sQry += "'" + PassYN + "',"; //경유지YN
                sQry += Amt + ","; //기본운임
                sQry += PassAmt + ","; //경유지운임
                sQry += TotAmt + ","; //총운임
                sQry += "'" + Comments + "'"; //비고

                RecordSet01.DoQuery(sQry);
                PSH_Globals.SBO_Application.StatusBar.SetText("수정 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 데이터 등록
        /// </summary>
        /// <returns></returns>
        private bool PS_MM004_AddData()
        {
            bool returnValue = false;
            string sQry;
            string errMessage = string.Empty;
            int DocEntry;
            string ObjType; //참조문서타입
            int DocNo;
            string BPLId; //사업장
            string DocDate; //일자
            string ItmBsort; //분류코드
            string ItmBName; //분류명
            string LocCode; //운송지역코드
            string LocName; //운송지역명
            string Way; //편도/회로
            string SectionF; //운행구간출발
            string SectionT; //운행구간도착
            string Destin; //도착장소
            string TranCard; //운송업체
            string TranCode; //차량번호
            double Tonnage; //차종(톤수)
            string ItemName; //품명
            double Weight; //실중량
            double TWeight; //총중량
            string PassYN; //경유지 YN
            double Amt; //금액
            double PassAmt; //경유지금액
            double TotAmt; //총금액
            string Comments; //비고
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ObjType = oForm.Items.Item("ObjType").Specific.Value.ToString().Trim(); //참조문서Object
                DocNo = Convert.ToInt32(oForm.Items.Item("DocNo").Specific.Value == "" ? "0" : oForm.Items.Item("DocNo").Specific.Value); //참조문서번호
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(); //사업장
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim(); //일자
                ItmBsort = oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim(); //품목분류
                ItmBName = oForm.Items.Item("ItmBName").Specific.Value.ToString().Trim(); //품목분류명
                LocCode = oForm.Items.Item("LocCode").Specific.Value.ToString().Trim(); //운송지역
                LocName = oForm.Items.Item("LocName").Specific.Value.ToString().Trim(); //운송지역명
                Way = oForm.Items.Item("Way").Specific.Value.ToString().Trim(); //편도/회로
                SectionF = oForm.Items.Item("SectionF").Specific.Value.ToString().Trim(); //운행구간(출발)
                SectionT = oForm.Items.Item("SectionT").Specific.Value.ToString().Trim(); //운행구간(도착)
                Destin = oForm.Items.Item("Destin").Specific.Value.ToString().Trim(); //도착장소
                TranCard = oForm.Items.Item("TranCard").Specific.Value.ToString().Trim(); //운송업체
                TranCode = oForm.Items.Item("TranCode").Specific.Value.ToString().Trim(); //차량번호
                Tonnage = Convert.ToDouble(oForm.Items.Item("Tonnage").Specific.Value); //차종(톤수)
                ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim(); //품명
                Weight = Convert.ToDouble(oForm.Items.Item("Weight").Specific.Value); //실중량
                TWeight = Convert.ToDouble(oForm.Items.Item("TWeight").Specific.Value); //총중량
                PassYN = oForm.Items.Item("PassYN").Specific.Value.ToString().Trim(); //경유지YN
                Amt = Convert.ToDouble(oForm.Items.Item("Amt").Specific.Value); //금액
                PassAmt = Convert.ToDouble(oForm.Items.Item("PassAmt").Specific.Value); //경유지금액
                TotAmt = Convert.ToDouble(oForm.Items.Item("TotAmt").Specific.Value); //총금액
                Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim(); //비고

                //DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_MM004H]";
                RecordSet01.DoQuery(sQry);

                if (Convert.ToInt32(RecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    DocEntry = 1;
                }
                else
                {
                    DocEntry = Convert.ToInt32(RecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1;
                }

                sQry = "EXEC [PS_MM004_02] ";
                sQry += DocEntry + ","; //관리번호
                sQry += "'" + ObjType + "',"; //참조문서Object
                sQry += DocNo + ","; //참조문서번호
                sQry += "'" + BPLId + "',"; //사업장
                sQry += "'" + DocDate + "',"; //일자
                sQry += "'" + ItmBsort + "',"; //분류코드
                sQry += "'" + ItmBName + "',"; //분류명
                sQry += "'" + LocCode + "',"; //운송지역
                sQry += "'" + LocName + "',"; //운송지역명
                sQry += "'" + Way + "',"; //편도/회로
                sQry += "'" + SectionF + "',"; //운행구간출발
                sQry += "'" + SectionT + "',"; //운행구간도착
                sQry += "'" + Destin + "',"; //도착장소
                sQry += "'" + TranCard + "',"; //운송업체
                sQry += "'" + TranCode + "',"; //차량번호
                sQry += Tonnage + ","; //차종(톤수)
                sQry += "'" + ItemName + "',"; //품명
                sQry += Weight + ","; //실중량
                sQry += TWeight + ","; //총중량
                sQry += "'" + PassYN + "',"; //경유지YN
                sQry += Amt + ","; //금액
                sQry += PassAmt + ","; //경유지금액
                sQry += TotAmt + ","; //금액
                sQry += "'" + Comments + "'"; //비고

                RecordSet02.DoQuery(sQry);
                PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet02);
            }
    
            return returnValue;
        }

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns>true:필수입력사항 모두 입력, fasle:필수입력사항 입력되지 않음</returns>
        private bool PS_MM004_CheckDataValie()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim())) //사업장
                {
                    errMessage = "사업장은 필수사항입니다. 확인하세요.";
                    oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim())) //품목분류
                {
                    errMessage = "품목분류는 필수사항입니다. 확인하세요.";
                    oForm.Items.Item("ItmBsort").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("LocCode").Specific.Value.ToString().Trim())) //운송지역코드
                {
                    errMessage = "운송지역코드는 필수사항입니다. 확인하세요.";
                    oForm.Items.Item("LocCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Way").Specific.Value.ToString().Trim())) //편도/회로
                {
                    errMessage = "편도/회로는 필수사항입니다. 확인하세요.";
                    oForm.Items.Item("Way").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("TranCard").Specific.Value.ToString().Trim())) //운송업체
                {
                    errMessage = "운송업체는 필수사항입니다. 확인하세요.";
                    oForm.Items.Item("TranCard").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("TranCode").Specific.Value.ToString().Trim())) //차량번호
                {
                    errMessage = "차량번호는 필수사항입니다. 확인하세요.";
                    oForm.Items.Item("TranCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Tonnage").Specific.Value.ToString().Trim())) //차종(톤수)
                {
                    errMessage = "차종(톤수)은 필수사항입니다. 확인하세요.";
                    oForm.Items.Item("Tonnage").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ItemName").Specific.Value.ToString().Trim())) //품명
                {
                    errMessage = "품명은 필수사항입니다. 확인하세요.";
                    oForm.Items.Item("Tonnage").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
        /// 운송확인증 출력
        /// </summary>
        [STAThread]
        private void PS_MM004_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                WinTitle = "[PS_MM004] 운송확인증";
                ReportName = "PS_MM004_10.rpt";
                //쿼리 : PS_MM004_10

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim())
                };

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
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
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "BtnAdd") //추가/확인 버튼클릭
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_MM004_CheckDataValie() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_MM004_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PS_MM004_SetForm("A");
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PS_MM004_LoadCaption();
                            PS_MM004_MTX01();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM004_CheckDataValie() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_MM004_UpdateData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PS_MM004_SetForm("U");
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PS_MM004_LoadCaption();
                            PS_MM004_MTX01();
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch") //조회
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                        PS_MM004_LoadCaption();
                        PS_MM004_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnDelete") //삭제
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
                        {
                            PS_MM004_DeleteData();
                            PS_MM004_SetForm("D");
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PS_MM004_LoadCaption();
                            PS_MM004_MTX01();
                        }
                        else
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnPrint")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_MM004_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
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
                        if (pVal.ItemUID == "ItmBsort")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ItmBsort").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "LocCode") //운송지역
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("LocCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Tonnage") //차종(톤수)
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("Tonnage").Specific.Value))
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
            double TWeight;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Way")
                    {
                        TWeight = Convert.ToDouble(oForm.Items.Item("TWeight").Specific.Value);

                        if (!string.IsNullOrEmpty(oForm.Items.Item("LocCode").Specific.Value.ToString().Trim()) && TWeight > 0)
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
                                oForm.Items.Item("Amt").Specific.Value = oRecordSet01.Fields.Item(0).Value;
                            }
                            else if (oForm.Items.Item("Way").Specific.Value.ToString().Trim() == "20")
                            {
                                oForm.Items.Item("Amt").Specific.Value = Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) * 0.8);
                            }

                        }

                        if (oForm.Items.Item("Way").Specific.Value.ToString().Trim() == "10")
                        {
                            oForm.Items.Item("SectionF").Specific.Value = "";
                            oForm.Items.Item("SectionT").Specific.Value = "";
                            oForm.Items.Item("ItemName").Specific.Value = "";
                            oForm.Items.Item("Destin").Specific.Value = "";
                        }
                        else if (oForm.Items.Item("Way").Specific.Value.ToString().Trim() == "20")
                        {
                            oForm.Items.Item("ItmBsort").Specific.Value = "102";
                            oForm.Items.Item("LocCode").Specific.Value = "07";
                            oForm.Items.Item("SectionF").Specific.Value = "안강";
                            oForm.Items.Item("SectionT").Specific.Value = "창원";
                            oForm.Items.Item("ItemName").Specific.Value = "부품소재";
                            oForm.Items.Item("Destin").Specific.Value = "창원";
                        }
                    }
                    else if (pVal.ItemUID == "PassYN")
                    {
                        if (oForm.Items.Item("PassYN").Specific.Value == "Y")
                        {
                            //경유지(혼적) Y이면
                            sQry = "  Select    PassAmt = Isnull(U_PassAmt,0)";
                            sQry += " From      [@PS_MM003H] a";
                            sQry += "           Inner Join";
                            sQry += "           [@PS_MM003L] b";
                            sQry += "               On a.Code = b.Code";
                            sQry += " Where     a.U_BPLId =  '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
                            sQry += "           And b.U_LocCode = '" + oForm.Items.Item("LocCode").Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oForm.Items.Item("PassAmt").Specific.Value = oRecordSet01.Fields.Item(0).Value;
                        }
                        else
                        {
                            oForm.Items.Item("PassAmt").Specific.Value = 0;
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                            PS_MM004_RET(oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value);
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
                    if (pVal.ItemChanged == true)
                    {
                        PS_MM004_FlushToItemValue(pVal.ItemUID, 0, "");
                    }
                }
                else if (pVal.Before_Action == false)
                {
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM004H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM004L);
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
                            PS_MM004_SetForm("A");
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            BubbleEvent = false;
                            PS_MM004_LoadCaption();
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

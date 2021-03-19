using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 비근무일수현황
    /// </summary>
    internal class PS_PP750 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Grid mainGrid;
        private SAPbouiCOM.DataTable oDS_PS_PP750;
        //private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        //private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        //private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP750.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP750_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP750");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                CreateItems();
                SetComboBox();
                InitializeForm();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// <returns></returns>
        private void CreateItems()
        {
            try
            {
                mainGrid = oForm.Items.Item("mainGrid").Specific;
                oForm.DataSources.DataTables.Add("PS_PP750");

                mainGrid.DataTable = oForm.DataSources.DataTables.Item("PS_PP750");
                oDS_PS_PP750 = oForm.DataSources.DataTables.Item("PS_PP750");

                //수주일자(FR)
                oForm.DataSources.UserDataSources.Add("FrDocDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("FrDocDt").Specific.DataBind.SetBound(true, "", "FrDocDt");

                //수주일자(TO)
                oForm.DataSources.UserDataSources.Add("ToDocDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ToDocDt").Specific.DataBind.SetBound(true, "", "ToDocDt");

                //납기일자(FR)
                oForm.DataSources.UserDataSources.Add("FrDueDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("FrDueDt").Specific.DataBind.SetBound(true, "", "FrDueDt");

                //납기일자(TO)
                oForm.DataSources.UserDataSources.Add("ToDueDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ToDueDt").Specific.DataBind.SetBound(true, "", "ToDueDt");

                //생산담당(사번)
                oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

                //생산담당(성명)
                oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

                //작번(품목코드)
                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                //품목명
                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

                //품목규격
                oForm.DataSources.UserDataSources.Add("ItemSpec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("ItemSpec").Specific.DataBind.SetBound(true, "", "ItemSpec");

                //자체/외주
                oForm.DataSources.UserDataSources.Add("InOut", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("InOut").Specific.DataBind.SetBound(true, "", "InOut");

                //장비/공구
                oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

                //거래처구분
                oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

                //조회구분
                oForm.DataSources.UserDataSources.Add("SrchType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("SrchType").Specific.DataBind.SetBound(true, "", "SrchType");

                //생산미완료(체크박스)
                oForm.DataSources.UserDataSources.Add("CmpltYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("CmpltYN").Specific.DataBind.SetBound(true, "", "CmpltYN");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                //자체/외주
                oForm.Items.Item("InOut").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("InOut").Specific.ValidValues.Add("IN", "자체");
                oForm.Items.Item("InOut").Specific.ValidValues.Add("OUT", "외주");
                oForm.Items.Item("InOut").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //장비/공구
                oForm.Items.Item("ItemType").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("ItemType").Specific.ValidValues.Add("M", "장비");
                oForm.Items.Item("ItemType").Specific.ValidValues.Add("T", "공구");
                oForm.Items.Item("ItemType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //거래처구분
                oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", false, false);
                oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //조회구분
                oForm.Items.Item("SrchType").Specific.ValidValues.Add("1", "간략보기");
                oForm.Items.Item("SrchType").Specific.ValidValues.Add("2", "상세보기");
                oForm.Items.Item("SrchType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
        /// Form 초기 세팅
        /// </summary>
        private void InitializeForm()
        {
            oForm.DataSources.UserDataSources.Item("FrDocDt").Value = DateTime.Now.ToString("yyyy0101");
            oForm.DataSources.UserDataSources.Item("ToDocDt").Value = DateTime.Now.ToString("yyyy1231");
            oForm.DataSources.UserDataSources.Item("FrDueDt").Value = DateTime.Now.ToString("yyyy0101");
            oForm.DataSources.UserDataSources.Item("ToDueDt").Value = DateTime.Now.ToString("yyyy1231");
            oForm.Items.Item("CmpltYN").Width = 90;
            oForm.Items.Item("BtnPrint").Visible = false; //출력버튼 비활성
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        /// <returns></returns>
        private void SelectData()
        {
            string sQry;
            string frDocDt; //수주일자(FR)
            string toDocDt; //수주일자(TO)
            string frDueDt; //납기일자(FR)
            string toDueDt; //납기일자(TO)
            string cntcCode; //생산담당
            string itemCode; //작번
            string inOut; //자체/외주
            string itemType; //장비/공구
            string cardType; //거래처구분
            string srchType; //조회구분
            string cmpltYN; //생산미완료여부
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                oForm.Freeze(true);
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                frDocDt = oForm.Items.Item("FrDocDt").Specific.Value.ToString().Trim(); //수주일자(FR)
                toDocDt = oForm.Items.Item("ToDocDt").Specific.Value.ToString().Trim(); //수주일자(TO)
                frDueDt = oForm.Items.Item("FrDueDt").Specific.Value.ToString().Trim(); //납기일자(FR)
                toDueDt = oForm.Items.Item("ToDueDt").Specific.Value.ToString().Trim(); //납기일자(TO)
                cntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim(); //생산담당
                itemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim(); //작번
                inOut = oForm.Items.Item("InOut").Specific.Value.ToString().Trim(); //자체/외주
                itemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim(); //장비/공구
                cardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim(); //거래처구분
                srchType = oForm.Items.Item("SrchType").Specific.Value.ToString().Trim(); //조회구분
                cmpltYN = (oForm.Items.Item("CmpltYN").Specific.Checked ? "Y" : "N"); //생산미완료여부

                if (srchType == "1") //간략보기
                {
                    sQry = "EXEC [PS_PP750_01] ";
                    sQry += "'" + frDocDt + "',";
                    sQry += "'" + toDocDt + "',";
                    sQry += "'" + frDueDt + "',";
                    sQry += "'" + toDueDt + "',";
                    sQry += "'" + cntcCode + "',";
                    sQry += "'" + itemCode + "',";
                    sQry += "'" + inOut + "',";
                    sQry += "'" + itemType + "',";
                    sQry += "'" + cardType + "',";
                    sQry += "'" + cmpltYN + "'";
                }
                else //상세보기
                {
                    sQry = "EXEC [PS_PP750_01] ";
                    sQry += "'" + frDocDt + "',";
                    sQry += "'" + toDocDt + "',";
                    sQry += "'" + frDueDt + "',";
                    sQry += "'" + toDueDt + "',";
                    sQry += "'" + cntcCode + "',";
                    sQry += "'" + itemCode + "',";
                    sQry += "'" + inOut + "',";
                    sQry += "'" + itemType + "',";
                    sQry += "'" + cardType + "',";
                    sQry += "'" + cmpltYN + "'";
                }

                mainGrid.DataTable.Clear();
                oDS_PS_PP750.ExecuteQuery(sQry);

                mainGrid.Columns.Item(5).RightJustified = true; //수주수량(5)
                mainGrid.Columns.Item(6).RightJustified = true; //수주금액(6)
                mainGrid.Columns.Item(13).RightJustified = true; //작번등록횟수(13)
                mainGrid.Columns.Item(16).RightJustified = true; //작업지시등록횟수(16)
                mainGrid.Columns.Item(18).RightJustified = true; //작업일보횟수(18)
                mainGrid.Columns.Item(20).RightJustified = true; //구매요청수량(20)
                mainGrid.Columns.Item(21).RightJustified = true; //구매요청횟수(21)
                mainGrid.Columns.Item(23).RightJustified = true; //구매견적수량(23)
                mainGrid.Columns.Item(24).RightJustified = true; //구매견적횟수(24)
                mainGrid.Columns.Item(26).RightJustified = true; //구매품의수량(26)
                mainGrid.Columns.Item(27).RightJustified = true; //구매품의횟수(27)
                mainGrid.Columns.Item(29).RightJustified = true; //가입고수량(29)
                mainGrid.Columns.Item(30).RightJustified = true; //가입고횟수(30)
                mainGrid.Columns.Item(32).RightJustified = true; //검수입고수량(32)
                mainGrid.Columns.Item(33).RightJustified = true; //검수입고횟수(33)
                mainGrid.Columns.Item(35).RightJustified = true; //검사횟수(35)
                mainGrid.Columns.Item(37).RightJustified = true; //생산완료수량(37)
                mainGrid.Columns.Item(38).RightJustified = true; //생산완료횟수(38)
                mainGrid.Columns.Item(39).RightJustified = true; //생산잔량(수주-생산)(39)

                mainGrid.Columns.Item(9).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //납기일(9)
                mainGrid.Columns.Item(13).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //작번등록횟수(13)
                mainGrid.Columns.Item(16).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //작업지시등록횟수(16)
                mainGrid.Columns.Item(18).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //작업일보횟수(18)
                mainGrid.Columns.Item(21).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //구매요청횟수(21)
                mainGrid.Columns.Item(24).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //구매견적횟수(24)
                mainGrid.Columns.Item(27).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //구매품의횟수(27)
                mainGrid.Columns.Item(30).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //가입고횟수(30)
                mainGrid.Columns.Item(33).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //검수입고횟수(33)
                mainGrid.Columns.Item(35).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //검사횟수(35)
                mainGrid.Columns.Item(38).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //생산완료횟수(38)
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 리포트 출력
        /// </summary>
        [STAThread]
        private void PrintReport()
        {
            //string WinTitle = string.Empty;
            //string ReportName = string.Empty;

            //string CLTCOD = string.Empty;
            //string YYYY = string.Empty;
            //string MSTCOD = string.Empty;
            //string TeamCode = string.Empty;
            //string RspCode = string.Empty;
            //string ClsCode = string.Empty;


            //PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            //PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            //CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
            //YYYY = oForm.Items.Item("YYYY").Specific.Value.Trim();
            //MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.Trim();
            //TeamCode = oForm.Items.Item("TeamCode").Specific.Value.Trim();
            //RspCode = oForm.Items.Item("RspCode").Specific.Value.Trim();
            //ClsCode = oForm.Items.Item("ClsCode").Specific.Value.Trim();

            try
            {
                //    WinTitle = "[PH_PY580] 비근무일수현황";
                //    ReportName = "PS_PP750_01.rpt";

                //    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();//Parameter List
                //    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                //    //Formula
                //    //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'")));
                //    //dataPackFormula.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr.Substring(0, 4) + "-" + DocDateFr.Substring(4, 2) + "-" + DocDateFr.Substring(6, 2)));
                //    //dataPackFormula.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo.Substring(0, 4) + "-" + DocDateTo.Substring(4, 2) + "-" + DocDateTo.Substring(6, 2)));

                //    //Parameter
                //    dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                //    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                //    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                //    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                //    dataPackParameter.Add(new PSH_DataPackClass("@MSTCOD", MSTCOD));
                //    dataPackParameter.Add(new PSH_DataPackClass("@YYYY", YYYY));

                //    formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
        /// ResizeForm
        /// </summary>
        private void ResizeForm()
        {
            try
            {
                if (mainGrid.Columns.Count > 0)
                {
                    mainGrid.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Raise_FormItemEvent
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">이벤트 </param>
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
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    // Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                    break;
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    break;
                case SAPbouiCOM.BoEventTypes.et_Drag: //39
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
                    if (pVal.ItemUID == "BtnSearch")
                    {
                        SelectData();
                    }
                    if (pVal.ItemUID == "BtnPrint")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PrintReport);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        if (pVal.ItemUID == "ItemCode" || pVal.ItemUID == "CntcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "ItemCode":
                                oForm.DataSources.UserDataSources.Item("ItemName").Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", "");
                                oForm.DataSources.UserDataSources.Item("ItemSpec").Value = dataHelpClass.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", "");
                                break;
                            case "CntcCode":
                                oForm.DataSources.UserDataSources.Item("CntcName").Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", "");
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mainGrid);
                }
                else if (pVal.Before_Action == false)
                {   
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    ResizeForm();
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

        #region Raise_FormMenuEvent
        ///// <summary>
        ///// FormMenuEvent
        ///// </summary>
        ///// <param name="FormUID"></param>
        ///// <param name="pVal"></param>
        ///// <param name="BubbleEvent"></param>
        //public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        oForm.Freeze(true);

        //        if (pVal.BeforeAction == true)
        //        {
        //            switch (pVal.MenuUID)
        //            {
        //                case "1284": //취소
        //                    break;
        //                case "1286": //닫기
        //                    break;
        //                case "1293": //행삭제
        //                    break;
        //                case "1281": //찾기
        //                    break;
        //                case "1282": //추가
        //                    break;
        //                case "1288":
        //                case "1289":
        //                case "1290":
        //                case "1291": //레코드이동버튼
        //                    break;
        //                case "7169": //엑셀 내보내기
        //                    break;
        //            }
        //        }
        //        else if (pVal.BeforeAction == false)
        //        {
        //            switch (pVal.MenuUID)
        //            {
        //                case "1284": //취소
        //                    break;
        //                case "1286": //닫기
        //                    break;
        //                case "1293": //행삭제
        //                    break;
        //                case "1281": //찾기
        //                    break;
        //                case "1282": //추가
        //                    break;
        //                case "1288":
        //                case "1289":
        //                case "1290":
        //                case "1291": //레코드이동버튼
        //                    break;
        //                case "7169": //엑셀 내보내기
        //                    break;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //    }
        //}
        #endregion

        #region Raise_FormDataEvent
        ///// <summary>
        ///// FormDataEvent
        ///// </summary>
        ///// <param name="FormUID"></param>
        ///// <param name="BusinessObjectInfo"></param>
        ///// <param name="BubbleEvent"></param>
        //public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}
        #endregion

        #region Raise_RightClickEvent
        ///// <summary>
        ///// RightClickEvent
        ///// </summary>
        ///// <param name="FormUID"></param>
        ///// <param name="pVal"></param>
        ///// <param name="BubbleEvent"></param>
        //public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        if (pVal.BeforeAction == true)
        //        {
        //        }
        //        else if (pVal.BeforeAction == false)
        //        {
        //        }

        //        switch (pVal.ItemUID)
        //        {
        //            case "Mat01":
        //                if (pVal.Row > 0)
        //                {
        //                    oLastItemUID01 = pVal.ItemUID;
        //                    oLastColUID01 = pVal.ColUID;
        //                    oLastColRow01 = pVal.Row;
        //                }
        //                break;
        //            default:
        //                oLastItemUID01 = pVal.ItemUID;
        //                oLastColUID01 = "";
        //                oLastColRow01 = 0;
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}
        #endregion
    }
}

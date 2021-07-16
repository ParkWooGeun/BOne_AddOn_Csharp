using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 기간별 표준공수 대비 실동공수 조회(작번별 상세)
    /// </summary>
    internal class PS_PP990 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Grid oGrid;

        private string gBPLID;
        private string gTeamCode;
        private string gRspCode;
        private string gClsCode;
        private string gCardType;
        private string gItemType;
        private string gWCYN;
        private string gDateStd;
        private string gFrDt;
        private string gToDt;
        private string gItemCode;
        private string gCpCode;
        private string gCpName;
        private string gOrdGbn;

        /// <summary>
        /// LoadForm
        /// </summary>
        /// <param name="pBPLID"></param>
        /// <param name="pTeamCode"></param>
        /// <param name="pRspCode"></param>
        /// <param name="pClsCode"></param>
        /// <param name="pCardType"></param>
        /// <param name="pItemType"></param>
        /// <param name="pWCYN"></param>
        /// <param name="pDateStd"></param>
        /// <param name="pFrDt"></param>
        /// <param name="pToDt"></param>
        /// <param name="pItemCode"></param>
        /// <param name="pCpCode"></param>
        /// <param name="pCpName"></param>
        /// <param name="pOrdGbn"></param>
        public void LoadForm(string pBPLID = "", string pTeamCode = "", string pRspCode = "", string pClsCode = "", string pCardType = "", string pItemType = "", string pWCYN = "", string pDateStd = "", string pFrDt = "", string pToDt = "",
        string pItemCode = "", string pCpCode = "", string pCpName = "", string pOrdGbn = "")
        {
            int i;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP990.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP990_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP990");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);

                PS_PP990_CreateItems();
                PS_PP990_MTX01(pBPLID, pTeamCode, pRspCode, pClsCode, pCardType, pItemType, pWCYN, pDateStd, pFrDt, pToDt, pItemCode, pCpCode, pCpName, pOrdGbn);

                gBPLID = pBPLID;
                gTeamCode = pTeamCode;
                gRspCode = pRspCode;
                gClsCode = pClsCode;
                gCardType = pCardType;
                gItemType = pItemType;
                gWCYN = pWCYN;
                gDateStd = pDateStd;
                gFrDt = pFrDt;
                gToDt = pToDt;
                gItemCode = pItemCode;
                gCpCode = pCpCode;
                gCpName = pCpName;
                gOrdGbn = pOrdGbn;
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
        /// PS_PP990_CreateItems
        /// </summary>
        /// <returns></returns>
        private void PS_PP990_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oGrid = oForm.Items.Item("Grid01").Specific;

                //공정코드
                oForm.DataSources.UserDataSources.Add("CpCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CpCode").Specific.DataBind.SetBound(true, "", "CpCode");

                //공정명
                oForm.DataSources.UserDataSources.Add("CpName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("CpName").Specific.DataBind.SetBound(true, "", "CpName");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PS_PP990_MTX01
        /// </summary>
        /// <param name="pBPLID"></param>
        /// <param name="pTeamCode"></param>
        /// <param name="pRspCode"></param>
        /// <param name="pClsCode"></param>
        /// <param name="pCardType"></param>
        /// <param name="pItemType"></param>
        /// <param name="pWCYN"></param>
        /// <param name="pDateStd"></param>
        /// <param name="pFrDt"></param>
        /// <param name="pToDt"></param>
        /// <param name="pItemCode"></param>
        /// <param name="pCpCode"></param>
        /// <param name="pCpName"></param>
        /// <param name="pOrdGbn"></param>
        private void PS_PP990_MTX01(string pBPLID, string pTeamCode, string pRspCode, string pClsCode, string pCardType, string pItemType, string pWCYN, string pDateStd, string pFrDt, string pToDt,        string pItemCode, string pCpCode, string pCpName, string pOrdGbn)
        {
            string sQry;
            string errMessage = String.Empty;

            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);

                oForm.Items.Item("CpCode").Specific.VALUE = pCpCode;
                oForm.Items.Item("CpName").Specific.VALUE = pCpName;

                ProgressBar01.Text = "조회 중...";

                sQry = " EXEC [PS_PP990_01] '";
                sQry += pBPLID + "','";    //사업장
                sQry += pTeamCode + "','"; //팀
                sQry += pRspCode + "','";  //담당
                sQry += pClsCode + "','";  //반
                sQry += pCardType + "','"; //거래처구분
                sQry += pItemType + "','"; //품목구분
                sQry += pWCYN + "','";     //생산완료여부
                sQry += pDateStd + "','";  //일자기준
                sQry += pFrDt + "','";     //기간(Fr)
                sQry += pToDt + "','";     //기간(To)
                sQry += pItemCode + "','"; //품목코드(작번)
                sQry += pCpCode + "','";   //공정
                sQry += pOrdGbn + "'";     //작업구분

                oGrid.DataTable.Clear();

                oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry);
                oGrid.DataTable = oForm.DataSources.DataTables.Item("DataTable");

                oGrid.Columns.Item(6).RightJustified = true;
                oGrid.Columns.Item(7).RightJustified = true;
                oGrid.Columns.Item(10).RightJustified = true;
                oGrid.Columns.Item(11).RightJustified = true;

                if (oGrid.Rows.Count == 0)
                {
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                oGrid.AutoResizeColumns();
                oForm.Update();
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
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_PP990_Print_Report01
        /// </summary>
        [STAThread]
        private void PS_PP990_Print_Report01()
        {
            string WinTitle;
            string ReportName;

            string BPLID;     //사업장
            string TeamCode;  //팀
            string RspCode;   //담당
            string ClsCode;   //반
            string CardType;  //거래처구분
            string ItemType;  //품목구분
            string WCYN;      //생산완료여부
            string DateStd;   //일자기준
            string FrDt;      //기간(Fr)
            string ToDt;      //기간(To)
            string ItemCode;  //품목코드(작번)
            string CpCode;    //공정
            string OrdGbn;    //작업구분

            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                BPLID = gBPLID;
                TeamCode = gTeamCode;
                RspCode = gRspCode;
                ClsCode = gClsCode;
                CardType = gCardType;
                ItemType = gItemType;
                WCYN = gWCYN;
                DateStd = gDateStd;
                FrDt = gFrDt;
                ToDt = gToDt;
                ItemCode = gItemCode;
                CpCode = gCpCode;
                OrdGbn = gOrdGbn;

                WinTitle = "[PS_PP990] 레포트";
                ReportName = "PS_PP990_01.rpt";

                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                // Formula 수식필드

                // Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
                dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
                dataPackParameter.Add(new PSH_DataPackClass("@ItemType", ItemType));
                dataPackParameter.Add(new PSH_DataPackClass("@WCYN", WCYN));
                dataPackParameter.Add(new PSH_DataPackClass("@DateStd", DateStd));
                dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
                dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));
                dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
                dataPackParameter.Add(new PSH_DataPackClass("@CpCode", CpCode));
                dataPackParameter.Add(new PSH_DataPackClass("@OrdGbn", OrdGbn));

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
        /// Raise_EVENT_ITEM_PRESSED
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "BtnPrint")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_PP990_Print_Report01);
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
        }

        /// <summary>
        /// Raise_EVENT_FORM_UNLOAD
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        case "1285": //복원
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
                        case "1285": //복원
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
        }

        /// <summary>
        /// Raise_FormDataEvent
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
    }
}

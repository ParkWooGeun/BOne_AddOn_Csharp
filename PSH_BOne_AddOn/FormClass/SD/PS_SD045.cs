using System;
using System.Linq;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Collections.Generic;
using PSH_BOne_AddOn.Code;
using System.Timers;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 반품처리(이월) 
    /// </summary>
    internal class PS_SD045 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_SD045H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_SD045L; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_SD045B; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        //DI API용 정보 클래스

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD045.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_SD045_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_SD045");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_SD045_CreateItems();
                PS_SD045_SetDocEntry();
                PS_SD045_EnableFormItem();

            //    oForm.EnableMenu("1283", false); //삭제
            //    oForm.EnableMenu("1287", false); //복제
            //    oForm.EnableMenu("1286", false); //닫기
            //    oForm.EnableMenu("1284", false); //취소
            //    oForm.EnableMenu("1293", false); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>        
        private void PS_SD045_CreateItems()
        {
            try
            {
                oDS_PS_SD045H = oForm.DataSources.DBDataSources.Item("@PS_SD045H");
                oDS_PS_SD045L = oForm.DataSources.DBDataSources.Item("@PS_SD045L");
                oDS_PS_SD045B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01"); //조회용 Matrix

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat01.AutoResizeColumns();

                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat02.AutoResizeColumns();

                //시작일
                oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

                //종료일
                oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");

                //거래처코드
                oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

                //거래처명
                oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }


        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_SD045_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD045'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0") 
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
            
        /// <summary>
        /// 각 모드에 따른 아이템설정
        /// </summary>
        private void PS_SD045_EnableFormItem()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("FrDt").Enabled = true;
                    oForm.Items.Item("ToDt").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("FrDt").Enabled = true;
                    oForm.Items.Item("ToDt").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("FrDt").Enabled = true;
                    oForm.Items.Item("ToDt").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가    
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 Row추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_SD045_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_SD045L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD045L.Offset = oRow;
                oDS_PS_SD045L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
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
        /// 필수입력사항 체크(헤더)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD045_DeleteHeaderSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("FrDt").Specific.Value.ToString().Trim())) //|| string.IsNullOrEmpty(oDS_PS_SD045H.GetValue("U_ToDt", 0)))
                {
                    errMessage = "전기년월은 필수입력사항입니다. 확인하세요.";
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// 필수입력사항 체크(라인)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD045_DeleteMatrixSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();
                if (oMat01.VisualRowCount == 0) //라인
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }

            return returnValue;
        }

     

        /// <summary>
        /// 데이터 조회(상단 list 조회)
        /// </summary>
        private void LoadData01()
        {
            short i = 0;
            string sQry;
            string errCode = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                string CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                string FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
                string ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();

                sQry = "exec [PS_SD045_01] '" + CardCode + "','" + FrDt + "','"+ ToDt + "'";
                oRecordSet01.DoQuery(sQry);

                oForm.Freeze(true);

                oDS_PS_SD045B.Clear();
                oMat02.Clear();
                oMat02.FlushToDataSource();
                oMat02.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errCode = "조회 결과가 없습니다. 확인하세요";
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i !=0)
                    {
                        oDS_PS_SD045B.InsertRecord(i);
                    }

                    oDS_PS_SD045B.Offset = i;
                    oDS_PS_SD045B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_SD045B.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim()); 
                    oDS_PS_SD045B.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("U_DocDate").Value.ToString().Trim()); 
                    oDS_PS_SD045B.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("U_CntcName").Value.ToString().Trim()); 
                    oDS_PS_SD045B.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim()); 
                    oDS_PS_SD045B.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim());  
                    oDS_PS_SD045B.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("U_TranCard").Value.ToString().Trim());    
                    oDS_PS_SD045B.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("U_DCardCod").Value.ToString().Trim()); 
                    oDS_PS_SD045B.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("U_DCardNam").Value.ToString().Trim());   
                    oDS_PS_SD045B.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errCode != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errCode);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                ProgBar01.Stop();
                oForm.Freeze(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
            }
        }
        /// <summary>
        /// 데이터 조회 (List 중 선택된 데이터)
        /// </summary>
        private void LoadData02(string DocEntry)
        {
            string sQry;
            string YYYYMM;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                YYYYMM = oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim();

                sQry = "EXEC [PS_SD045_01] '" +  "','" + YYYYMM + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_SD045L.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다.확인하세요.";
                    throw new Exception();
                }

                for (int i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_SD045L.Size)
                    {
                        oDS_PS_SD045L.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_SD045L.Offset = i;
                    oDS_PS_SD045L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_SD045L.SetValue("U_ODLNDocB", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_DLN1LinB", i, oRecordSet01.Fields.Item("LineNum").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_CardCode", i, oRecordSet01.Fields.Item("CardCode").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_CardName", i, oRecordSet01.Fields.Item("CardName").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_ItemCode", i, oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_ItemName", i, oRecordSet01.Fields.Item("Dscription").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_WhsCode", i, oRecordSet01.Fields.Item("WhsCode").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_BatchNum", i, oRecordSet01.Fields.Item("BatchNum").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_BatchQty", i, oRecordSet01.Fields.Item("BatchQty").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_Quantity", i, oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_OpenQty", i, oRecordSet01.Fields.Item("OpenQty").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_Price", i, oRecordSet01.Fields.Item("Price").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_LinTotal", i, oRecordSet01.Fields.Item("LineTotal").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_Qty", i, oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_BaseType", i, oRecordSet01.Fields.Item("U_BaseType").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_BaseDoc", i, oRecordSet01.Fields.Item("U_BaseEntry").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_BaseLine", i, oRecordSet01.Fields.Item("U_BaseLine").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_ULineNum", i, oRecordSet01.Fields.Item("U_LineNum").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_ORDRDoc", i, oRecordSet01.Fields.Item("BaseEntry").Value.ToString().Trim());
                    oDS_PS_SD045L.SetValue("U_RDR1Line", i, oRecordSet01.Fields.Item("BaseLine").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
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
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD045_DeleteHeaderSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD045_DeleteHeaderSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            LoadData01();
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        if (pVal.ItemUID == "1")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                            {
                            }
                        }
                        else if (pVal.ItemUID == "Btn01")
                        {
                            if (PS_SD045_DeleteHeaderSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            LoadData01();
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
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string ColReg01;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat02" && pVal.ColUID == "ColReg01")
                    {
                        ColReg01 = oMat02.Columns.Item("ColReg01").Cells.Item(pVal.Row).Specific.Value;
                        PS_SD040 pS_SD040 = new PS_SD040();
                        pS_SD040.LoadForm(ColReg01);
                        pS_SD040 = null;
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
                                sQry = "SELECT CardFName FROM OCRD WHERE CardCode =  '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item("CardFName").Value.ToString().Trim();
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
                    if (pVal.ItemUID == "CardCode" || pVal.ItemUID == "CardName")
                    {
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_SD045H", "U_CardCode,U_CardName", "", 0, "", "", "");
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "CardCode")
                    {
                        if (pVal.CharPressed == 9) //탭을 눌렀을 경우만
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
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
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string check = string.Empty;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat02.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat02.FlushToDataSource();
                        }
                        else
                        {
                            LoadData02(oMat02.Columns.Item("ColReg01").Cells.Item(pVal.Row).Specific.Value); //매트릭스
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
                    PS_SD045_AddMatrixRow(oMat01.RowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD045H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD045L);
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
                            PS_SD045_EnableFormItem();
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
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            //oForm.Freeze(true);
                            //if (oMat01.RowCount != oMat01.VisualRowCount)
                            //{
                            //    for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            //    {
                            //        oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
                            //    }

                            //    oMat01.FlushToDataSource();
                            //    oDS_PS_SD045L.RemoveRecord(oDS_PS_SD045L.Size - 1);
                            //    oMat01.Clear();
                            //    oMat01.LoadFromDataSource();

                            //    if (!string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(oMat01.RowCount).Specific.Value))
                            //    {
                            //        PS_SD045_AddMatrixRow(oMat01.RowCount, false);
                            //    }
                            //}
                            //oForm.Freeze(false);
                            break;
                        case "1281": //찾기
                            PS_SD045_EnableFormItem();
                            break;
                        case "1282": //추가
                            PS_SD045_EnableFormItem();
                            PS_SD045_SetDocEntry();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_SD045_EnableFormItem();
                            break;
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
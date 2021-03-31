using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 날짜 변경 요청
    /// </summary>
    internal class PS_DateChange : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PS_DateChangeA;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_DateChange.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PS_DateChange_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PS_DateChange");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_DateChange_CreateItems();
                PS_DateChange_FormItemEnabled();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                //oForm.Visible = true;
                //oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_DateChange_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PS_DateChange");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PS_DateChange");
                oDS_PS_DateChangeA = oForm.DataSources.DataTables.Item("PS_DateChange");

                // BPLId
                sQry = " SELECT T2.Code ,T2.Name";
                sQry += " From [@PH_PY000B] T0 INNER JOIN [@PH_PY000A] T1 ON T0.Code = T1.Code";
                sQry += " INNER JOIN [@PH_PY005A] T2 ON T0.U_Value = T2.Code";
                sQry += " WHERE T1.Code = 'CLTCOD' AND T0.U_UserCode = '" + PSH_Globals.oCompany.UserName + "'";
                sQry += " GROUP BY T2.Code ,T2.Name ORDER BY T2.Code";

                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("BPLId").Specific, "N");
                oForm.Items.Item("BPLId").DisplayDesc = true;
                oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                // ObjectCode
                sQry = "select Code, Name from [@PS_SY005H] where Left(Code ,1)  <>'S'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ObjectCode").Specific, "Y");
                oForm.Items.Item("ObjectCode").DisplayDesc = true;
                oForm.Items.Item("ObjectCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // CreateUser
                oForm.DataSources.UserDataSources.Add("CreateUser", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CreateUser").Specific.DataBind.SetBound(true, "", "CreateUser");
                oForm.Items.Item("CreateUser").Specific.Value = PSH_Globals.oCompany.UserName;

                // CreateUseV
                oForm.DataSources.UserDataSources.Add("CreateUseV", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CreateUseV").Specific.DataBind.SetBound(true, "", "CreateUseV");
                oForm.Items.Item("CreateUseV").Specific.Value = dataHelpClass.Get_ReData("U_Name", "USER_CODE", "OUSR", "'" + PSH_Globals.oCompany.UserName + "'", "");
                
                // Grantor 승인자
                oForm.DataSources.UserDataSources.Add("Grantor", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Grantor").Specific.DataBind.SetBound(true, "", "Grantor");

                // GrantorV
                oForm.DataSources.UserDataSources.Add("GrantorV", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("GrantorV").Specific.DataBind.SetBound(true, "", "GrantorV");

                // CreateDate
                oForm.DataSources.UserDataSources.Add("CreateDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("CreateDate").Specific.DataBind.SetBound(true, "", "CreateDate");
                oForm.Items.Item("CreateDate").Specific.Value =  DateTime.Now.ToString("yyyyMMdd");

                // DocEntry
                oForm.DataSources.UserDataSources.Add("DocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("DocEntry").Specific.DataBind.SetBound(true, "", "DocEntry");

                // LineId
                oForm.DataSources.UserDataSources.Add("LineId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
                oForm.Items.Item("LineId").Specific.DataBind.SetBound(true, "", "LineId");

                // DocDate
                oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");

                // DueDate
                oForm.DataSources.UserDataSources.Add("DueDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("DueDate").Specific.DataBind.SetBound(true, "", "DueDate");

                // TaxDate
                oForm.DataSources.UserDataSources.Add("TaxDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("TaxDate").Specific.DataBind.SetBound(true, "", "TaxDate");

                // Comments
                oForm.DataSources.UserDataSources.Add("Comments", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("Comments").Specific.DataBind.SetBound(true, "", "Comments");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PS_DateChange_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PS_DateChange_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", false); // 문서찾기
                    oForm.EnableMenu("1282", true); // 문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", true); // 문서찾기
                    oForm.EnableMenu("1282", false); // 문서추가
                    dataHelpClass.CLTCOD_Select(oForm, "BPLID", true);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); // 문서찾기
                    oForm.EnableMenu("1282", true); // 문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPS_DateChange_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_DateChange_MTX01
        /// </summary>
        private void PS_DateChange_MTX01()
        {
            int iRow;
            string sQry;
            string BPLId;
            string CreateUser;
            string ObjectCode;

            try
            {
                oForm.Freeze(true);
                oForm.Items.Item("DocEntry").Specific.Value = "";
                oForm.Items.Item("LineId").Specific.Value = "";
                oForm.Items.Item("DocDate").Specific.Value = "";
                oForm.Items.Item("DueDate").Specific.Value = "";
                oForm.Items.Item("TaxDate").Specific.Value = "";
                oForm.Items.Item("Comments").Specific.Value = "";
                oForm.Items.Item("OKYN").Specific.Value = "";

                oForm.Items.Item("Grantor").Specific.Value = "";
                oForm.Items.Item("GrantorV").Specific.Value = "";

                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                CreateUser = oForm.Items.Item("CreateUser").Specific.Value.ToString().Trim();
                ObjectCode = oForm.Items.Item("ObjectCode").Specific.Value.ToString().Trim();

                sQry = "EXEC PS_DateChange_01 '" + BPLId + "', '" + CreateUser + "', '" + ObjectCode + "'";
                oDS_PS_DateChangeA.ExecuteQuery(sQry);
                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PS_DateChange_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_DateChange_MTX02
        /// </summary>
        private void PS_DateChange_MTX02(int oRow)
        {
            try
            {
                oForm.Freeze(true);
                oForm.Items.Item("CreateDate").Specific.Value = oDS_PS_DateChangeA.Columns.Item("등록일").Cells.Item(oRow).Value;

                oForm.Items.Item("DocEntry").Specific.Value = oDS_PS_DateChangeA.Columns.Item("문서번호").Cells.Item(oRow).Value;
                oForm.Items.Item("LineId").Specific.Value = oDS_PS_DateChangeA.Columns.Item("행번호").Cells.Item(oRow).Value;
                oForm.Items.Item("DocDate").Specific.Value = oDS_PS_DateChangeA.Columns.Item("전기일").Cells.Item(oRow).Value;
                oForm.Items.Item("DueDate").Specific.Value = oDS_PS_DateChangeA.Columns.Item("만기일").Cells.Item(oRow).Value;
                oForm.Items.Item("TaxDate").Specific.Value = oDS_PS_DateChangeA.Columns.Item("증빙일").Cells.Item(oRow).Value;
                oForm.Items.Item("Comments").Specific.Value = oDS_PS_DateChangeA.Columns.Item("관련근거").Cells.Item(oRow).Value;
                oForm.Items.Item("OKYN").Specific.Value = oDS_PS_DateChangeA.Columns.Item("처리상태").Cells.Item(oRow).Value;

                oForm.Items.Item("Grantor").Specific.Value = oDS_PS_DateChangeA.Columns.Item("승인자").Cells.Item(oRow).Value;
                oForm.Items.Item("GrantorV").Specific.Value = oDS_PS_DateChangeA.Columns.Item("승인자명").Cells.Item(oRow).Value;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PS_DateChange_MTX02_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_DateChange_SAVE
        /// </summary>
        private void PS_DateChange_SAVE()
        {
            // 데이타 저장
            int ErrNum = 0;
            string sQry;
            string BPLId;
            string ObjectCode;
            string CreateUser;
            string CreateUseV;
            string Grantor;
            string GrantorV;
            string CreateDate;
            string DocEntry;
            string LineId;
            string DocDate;
            string DueDate;
            string TaxDate;
            string Comments;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (PSH_Globals.SBO_Application.MessageBox("저장하시겠습니까?", 2, "Yes", "No") == 2)
                {
                    ErrNum = 7;
                    throw new Exception();
                }

                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                ObjectCode = oForm.Items.Item("ObjectCode").Specific.Value.ToString().Trim();
                CreateUser = oForm.Items.Item("CreateUser").Specific.Value.ToString().Trim();
                CreateUseV = oForm.Items.Item("CreateUseV").Specific.Value.ToString().Trim();
                Grantor = oForm.Items.Item("Grantor").Specific.Value.ToString().Trim();
                GrantorV = oForm.Items.Item("GrantorV").Specific.Value.ToString().Trim();
                CreateDate = oForm.Items.Item("CreateDate").Specific.Value.ToString().Trim();
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                LineId = oForm.Items.Item("LineId").Specific.Value.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                DueDate = oForm.Items.Item("DueDate").Specific.Value.ToString().Trim();
                TaxDate = oForm.Items.Item("TaxDate").Specific.Value.ToString().Trim();
                Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim();

                if (BPLId == "")
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (ObjectCode == "")
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                if (Grantor == "")
                {
                    ErrNum = 3;
                    throw new Exception();
                }
                if (DocEntry == "")
                {
                    ErrNum = 4;
                    throw new Exception();
                }
                if (DocDate + DueDate + TaxDate == "")
                {
                    ErrNum = 5;
                    throw new Exception();
                }
                if (Comments == "")
                {
                    ErrNum = 6;
                    throw new Exception();
                }

                sQry = "EXEC PS_DateChange_02 '" + BPLId + "', '" + ObjectCode + "', '" + CreateUser + "', '";
                sQry += CreateUseV + "', '" + Grantor + "', '" + GrantorV + "', '" + CreateDate + "', '" + DocEntry + "', '" + LineId + "', '" + DocDate + "', '" + DueDate + "', '";
                sQry += TaxDate + "', '" + Comments + "'";
                //oDS_PS_DateChangeA.ExecuteQuery(sQry);
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item("returnV").Value == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("수정완료");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox("신규등록");
                }

                PS_DateChange_MTX01();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장은 필수입니다.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("구분은 필수입니다.");
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.MessageBox("승인자는 필수입니다.");
                }
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.MessageBox("문서번호는 필수입니다.");
                }
                else if (ErrNum == 5)
                {
                    PSH_Globals.SBO_Application.MessageBox("날짜입력은 필수입니다.");
                }
                else if (ErrNum == 6)
                {
                    PSH_Globals.SBO_Application.MessageBox("관련근거 입력은 필수입니다.");
                }
                else if (ErrNum == 7)
                {
                    PSH_Globals.SBO_Application.MessageBox("저장을 취소하셨습니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PS_DateChange_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_DateChange_Delete
        /// </summary>
        private void PS_DateChange_Delete()
        {
            string sQry;
            int ErrNum = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                if (PSH_Globals.SBO_Application.MessageBox("삭제하시겠습니까?", 2, "Yes", "No") == 2)
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (oForm.Items.Item("OKYN").Specific.Value.Trim() != "N")
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                sQry = "delete from PSH_DateChange where ObjectCode = '" + oForm.Items.Item("ObjectCode").Specific.Value + "'";
                sQry += " and DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                sQry += " and LineId = '" + oForm.Items.Item("LineId").Specific.Value + "'";
                sQry += " and OKYN = 'N'";
                oRecordSet.DoQuery(sQry);

                PS_DateChange_MTX01();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("삭제가 취소되었습니다.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("처리상태가 N인것만 삭제가능합니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PS_DateChange_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }
        
        /// <summary>
        /// Form Item Event
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
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
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                    if (pVal.ItemUID == "1")
                    {
                    }

                    if (pVal.ItemUID == "Btn_Find")
                    {
                        PS_DateChange_MTX01();
                    }

                    if (pVal.ItemUID == "Btn01")
                    {
                        PS_DateChange_SAVE();
                    }

                    if (pVal.ItemUID == "Btn_del")
                    {
                        PS_DateChange_Delete();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.ItemUID)
                    {
                        case "1":
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PS_DateChange_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PS_DateChange_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PS_DateChange_FormItemEnabled();
                                }
                            }
                            break;
                    }
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
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        default:
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = "";
                            oLastColRow = 0;
                            break;
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        // 사업장(헤더)
                        switch (pVal.ItemUID)
                        {
                            case "Modual":
                            case "No":
                            case "Sub1":
                            case "Sub2":
                            case "ObjectCode":
                                if (oForm.Items.Item("ObjectCode").Specific.Value.ToString().Trim() =="31")
                                {
                                    oForm.Items.Item("DocDate").Enabled = false;
                                    oForm.Items.Item("DueDate").Enabled = false;
                                    oForm.Items.Item("LineId").Enabled = true;
                                }
                                else
                                {
                                    oForm.Items.Item("DocDate").Enabled = true;
                                    oForm.Items.Item("DueDate").Enabled = true;
                                    oForm.Items.Item("LineId").Enabled = false;
                                }
                                PS_DateChange_MTX01();
                                break;
                        }
                    }
                }
                PS_DateChange_FormItemEnabled();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "Grantor":
                                oForm.Items.Item("GrantorV").Specific.Value = dataHelpClass.Get_ReData("U_Name", "USER_CODE", "OUSR", "'" + oForm.Items.Item("Grantor").Specific.Value + "'", "");
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
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
                            if (pVal.Row >= 0)
                            {
                                switch (pVal.ItemUID)
                                {
                                    case "Grid01":
                                        PS_DateChange_MTX02(pVal.Row);
                                        break;
                                }
                            }
                            break;
                    }

                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        default:
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = "";
                            oLastColRow = 0;
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                }
                else if (pVal.Before_Action == false)
                {
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_DateChangeA);
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
                            if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
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
                            PS_DateChange_FormItemEnabled();
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PS_DateChange_FormItemEnabled();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
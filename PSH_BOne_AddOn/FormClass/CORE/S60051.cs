using System;
using SAPbouiCOM;

namespace PSH_BOne_AddOn.Core
{
    /// <summary>
    /// 자금관리>어음관리-어음관리
    /// </summary>
    internal class S60051 : PSH_BaseClass
    {
        private SAPbouiCOM.Matrix oMat;
        private int oMatRow;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="formUID"></param>
        public override void LoadForm(string formUID)
        {
            try
            {
                oForm = PSH_Globals.SBO_Application.Forms.Item(formUID);
                oForm.Freeze(true);
                SubMain.Add_Forms(this, formUID, "S60051");
                S60051_CreateItems();
                oMat = oForm.Items.Item("5").Specific;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// S60051_CreateItems
        /// </summary>
        private void S60051_CreateItems()
        {
            SAPbouiCOM.Item newItem = null;

            try
            {
                newItem = oForm.Items.Add("AddonText", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                newItem.Top = oForm.Items.Item("1").Top - 12;
                newItem.Left = oForm.Items.Item("1").Left;
                newItem.Height = 12;
                newItem.Width = 70;
                newItem.FontSize = 10;
                newItem.Specific.Caption = "Addon running";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(newItem);
            }
        }

        /// <summary>
        /// S60051_Create_oJournalEntries
        /// </summary>
        /// <returns></returns>
        private bool S60051_Create_oJournalEntries()
        {
            bool ReturnValue = false;
            int i;
            int j;
            int RetVal;
            string VTransId;
            string vBoeKey;
            string vBPLId;
            int ErrCode = 0;
            string errCode = string.Empty;
            string ErrMsg = string.Empty;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.JournalEntries oJournal = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                i = 0;
                //var _with1 = f_oJournalEntries;

                // Header
                oJournal.ReferenceDate = DateTime.ParseExact(oForm.Items.Item("55").Specific.Value.ToString().Trim(), "yyyyMMdd", null);
                oJournal.DueDate = DateTime.ParseExact(oForm.Items.Item("55").Specific.Value.ToString().Trim(), "yyyyMMdd", null);
                oJournal.TaxDate = DateTime.ParseExact(oForm.Items.Item("61").Specific.Value.ToString().Trim(), "yyyyMMdd", null);
                oJournal.Memo = "추심에서 부도어음이동";

                // Line
                for (j = 1; j <= oMat.VisualRowCount; j++)
                {
                    if (Convert.ToBoolean(oMat.Columns.Item("9").Cells.Item(j).Specific.Checked) == true)
                    {
                        sQry = "select BoeKey from OBOE where BoeType = 'I' and BoeNum = '" + oMat.Columns.Item("7").Cells.Item(j).Specific.Value.ToString().Trim() + "'";
                        oRecordSet.DoQuery(sQry);
                        vBoeKey = oRecordSet.Fields.Item("BoeKey").Value.ToString().Trim();

                        sQry = "select U_BPLId from ORCT where BoeAbs = '" + vBoeKey + "'";
                        oRecordSet.DoQuery(sQry);
                        vBPLId = oRecordSet.Fields.Item("U_BPLId").Value.ToString().Trim();
                        //전표헤더 사업장
                        oJournal.UserFields.Fields.Item("U_BPLId").Value = vBPLId;

                        //차변(Debit)--------------------------------------------------------
                        oJournal.Lines.Add();
                        oJournal.Lines.SetCurrentLine(i);

                        oJournal.Lines.ShortName = oMat.Columns.Item("28").Cells.Item(j).Specific.Value.ToString().Trim();
                        oJournal.Lines.ControlAccount = "11104070";
                        //부도어음
                        oJournal.Lines.Debit = Convert.ToDouble(oMat.Columns.Item("2").Cells.Item(j).Specific.Value.ToString().Trim());
                        oJournal.Lines.Reference1 = vBoeKey;
                        oJournal.Lines.LineMemo = "어음관리 번호(" + oMat.Columns.Item("7").Cells.Item(j).Specific.Value.ToString().Trim() + ") : 추심에서 부도이동";
                        oJournal.Lines.UserFields.Fields.Item("U_BPLId").Value = vBPLId;
                        i += 1;

                        //대변(Credit)
                        oJournal.Lines.Add();
                        oJournal.Lines.SetCurrentLine(i);
                        oJournal.Lines.ShortName = oMat.Columns.Item("28").Cells.Item(j).Specific.Value.ToString().Trim();
                        oJournal.Lines.ControlAccount = "11104060";
                        //받을어음
                        oJournal.Lines.Credit = Convert.ToDouble(oMat.Columns.Item("2").Cells.Item(j).Specific.Value.ToString().Trim());
                        oJournal.Lines.Reference1 = vBoeKey;
                        oJournal.Lines.LineMemo = "어음관리 번호(" + oMat.Columns.Item("7").Cells.Item(j).Specific.Value.ToString().Trim() + ") : 추심에서 부도이동";
                        oJournal.Lines.UserFields.Fields.Item("U_BPLId").Value = vBPLId;
                        i += 1;
                    }
                }
                RetVal = oJournal.Add(); // 완료

                if (0 != RetVal)
                {
                    PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
                    errCode = "1";
                    throw new Exception();
                }
                else
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out VTransId);

                    for (j = 1; j <= oMat.VisualRowCount; j++)
                    {
                        if (Convert.ToBoolean(oMat.Columns.Item("9").Cells.Item(j).Specific.Checked) == true)
                        {
                            // OBOE 내용 Upate
                            sQry = " update OBOE set BoeStatus='F'";    //F : 부도어음
                            sQry += " where BoeType='I'"; // I : 받을어음
                            sQry += " and BoeNum='" + oMat.Columns.Item("7").Cells.Item(j).Specific.Value.ToString().Trim() + "'";
                            oRecordSet.DoQuery(sQry);

                            // 정보저장 Insert
                            sQry = "select BoeKey from OBOE where BoeType = 'I' and BoeNum = '" + oMat.Columns.Item("7").Cells.Item(j).Specific.Value.ToString().Trim() + "'";
                            oRecordSet.DoQuery(sQry);

                            vBoeKey = oRecordSet.Fields.Item("BoeKey").Value.ToString().Trim();

                            sQry = "insert into Z60051 values('" + VTransId + "','" + vBoeKey + "','" + oMat.Columns.Item("7").Cells.Item(j).Specific.Value.ToString().Trim() + "')";
                            oRecordSet.DoQuery(sQry);
                        }
                    }

                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

                ReturnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + ErrCode + "]" + ErrMsg);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournal);
            }
            return ReturnValue;
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
                //	Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //	break;
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
                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //	Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //	Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                //	Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //	break;
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
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            //부도일 경우
                            if (oForm.Items.Item("4").Specific.Value.ToString().Trim() == "F")
                            {
                                if (S60051_Create_oJournalEntries() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                else
                                {
                                    PSH_Globals.SBO_Application.MessageBox("부도어음으로 이동이 완료되었습니다.");
                                }
                                oForm.Items.Item("38").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                BubbleEvent = false;
                            }
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
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "38")
                    {
                        oForm.Items.Item("4").Specific.ValidValues.Add("F", "부도");
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Raise_RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="eventInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
            try
            {
                if (eventInfo.BeforeAction == true)
                {
                    if (eventInfo.ItemUID == "76")
                    {
                        if (eventInfo.Row > 0)
                        {
                            oMatRow = eventInfo.Row;
                        }
                    }
                }
                else if (eventInfo.BeforeAction == false)
                {
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
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1283": //삭제
                            break;
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "1293": //행삭제
                            break;
                        case "7169": //엑셀 내보내기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1287": // 복제
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "1293": //행삭제
                            break;
                        case "7169": //엑셀 내보내기
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
        }
    }
}

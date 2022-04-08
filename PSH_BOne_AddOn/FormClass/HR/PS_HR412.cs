using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 전문직 평가 그룹핑
    /// </summary>
    internal class PS_HR412 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_HR412L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_HR412.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_HR412_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_HR412");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_HR412_CreateItems();
                PS_HR412_ComboBox_Setting();
                PS_HR412_Initialization();
                PS_HR412_SetDocument(oFromDocEntry01);

                oForm.EnableMenu("1281", false);
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
        private void PS_HR412_CreateItems()
        {
            try
            {
                oDS_PS_HR412L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

                //년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

                //차수
                oForm.DataSources.UserDataSources.Add("Number", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Number").Specific.DataBind.SetBound(true, "", "Number");

                //평가권한
                oForm.DataSources.UserDataSources.Add("Evaluate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Evaluate").Specific.DataBind.SetBound(true, "", "Evaluate");

                //사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                //성명
                oForm.DataSources.UserDataSources.Add("FULLNAME", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FULLNAME").Specific.DataBind.SetBound(true, "", "FULLNAME");

                //직급명
                oForm.DataSources.UserDataSources.Add("JigNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("JigNm").Specific.DataBind.SetBound(true, "", "JigNm");
                
                //호칭
                oForm.DataSources.UserDataSources.Add("CallName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CallName").Specific.DataBind.SetBound(true, "", "CallName");
               
                //직책
                oForm.DataSources.UserDataSources.Add("JicNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("JicNm").Specific.DataBind.SetBound(true, "", "JicNm");

                //팀코드
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

                //팀명
                oForm.DataSources.UserDataSources.Add("TeamNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("TeamNm").Specific.DataBind.SetBound(true, "", "TeamNm");

                //담당코드
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");

                //담당명
                oForm.DataSources.UserDataSources.Add("RspNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RspNm").Specific.DataBind.SetBound(true, "", "RspNm");

                //반코드
                oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

                //반명
                oForm.DataSources.UserDataSources.Add("ClsNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ClsNm").Specific.DataBind.SetBound(true, "", "ClsNm");

                //평가대상자그룹
                oForm.DataSources.UserDataSources.Add("Group", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Group").Specific.DataBind.SetBound(true, "", "Group");

                //평가대상자그룹(조회용)
                oForm.DataSources.UserDataSources.Add("SGroup", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("SGroup").Specific.DataBind.SetBound(true, "", "SGroup");

                oForm.Items.Item("Mat01").Enabled = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_HR412_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);

                // 평가차수
                oForm.Items.Item("Number").Specific.ValidValues.Add("1", "1차");
                oForm.Items.Item("Number").Specific.ValidValues.Add("2", "2차");

                //평가권한
                oForm.Items.Item("Evaluate").Specific.ValidValues.Add("1", "1차평가자");
                oForm.Items.Item("Evaluate").Specific.ValidValues.Add("2", "2차평가자");
                oForm.Items.Item("Evaluate").Specific.ValidValues.Add("3", "종합평가자");

                //평가대상자그룹
                oForm.Items.Item("Group").Specific.ValidValues.Add("1", "반장");
                oForm.Items.Item("Group").Specific.ValidValues.Add("2", "사원");

                //평가대상자그룹(조회용)
                oForm.Items.Item("SGroup").Specific.ValidValues.Add("1", "반장");
                oForm.Items.Item("SGroup").Specific.ValidValues.Add("2", "사원");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Initialization
        /// </summary>
        private void PS_HR412_Initialization()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFromDocEntry01">DocEntry</param>
        private void PS_HR412_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PS_HR412_AddMatrixRow(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    oForm.Items.Item("DocEntry").Specific.Value = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_HR412_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_HR412_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_HR412L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_HR412L.Offset = oRow;
                oDS_PS_HR412L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        /// PS_HR412_MTX01
        /// </summary>
        private void PS_HR412_MTX01()
        {
            string errMessage = string.Empty; 
            int i;
            string sQry; 
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            string Param06;
            string Param07;
            string Param08;
            string Param09;
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(); //사업장
                Param02 = oForm.Items.Item("Year").Specific.Value.ToString().Trim(); //년도
                Param03 = oForm.Items.Item("Number").Specific.Value.ToString().Trim(); //평가차수 1차, 2차평가
                Param04 = oForm.Items.Item("Evaluate").Specific.Value.ToString().Trim(); //평가권한 1차,2차,3차권한
                Param05 = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim(); //팀코드
                Param06 = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim(); //담당코드
                Param07 = oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim(); //반코드
                Param08 = oForm.Items.Item("SGroup").Specific.Value.ToString().Trim(); //그룹 1:반장, 2:사원
                Param09 = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim(); //평가자사번

                ProgressBar01.Text = "조회시작!";

                sQry = "EXEC PS_HR412_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "', '" + Param08 + "', '" + Param09 + "'";
                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    oForm.Items.Item("Mat01").Enabled = false;
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                else
                {
                    oForm.Items.Item("Mat01").Enabled = true;
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_HR412L.InsertRecord(i);
                    }
                    oDS_PS_HR412L.Offset = i;
                    oDS_PS_HR412L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_HR412L.SetValue("U_ColReg01", i, Convert.ToString(false));
                    oDS_PS_HR412L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("MSTCOD").Value); //사번
                    oDS_PS_HR412L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("FULLNAME").Value); //성명
                    oDS_PS_HR412L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("TeamNm").Value); //팀명
                    oDS_PS_HR412L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("RspNm").Value); //담당명
                    oDS_PS_HR412L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("JigNm").Value); //직급
                    oDS_PS_HR412L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("CallName").Value); //호칭
                    oDS_PS_HR412L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("JicNm").Value); //직책
                    oDS_PS_HR412L.SetValue("U_ColReg15", i, oRecordSet.Fields.Item("PeakYN").Value); //피크사원여
                    oDS_PS_HR412L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("MSTCOD1").Value); //1차평가자사번
                    oDS_PS_HR412L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("FULLNAM1").Value); //1차평가자성명
                    oDS_PS_HR412L.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("MSTCOD2").Value); //2차평가자사번
                    oDS_PS_HR412L.SetValue("U_ColReg12", i, oRecordSet.Fields.Item("FULLNAM2").Value); //2차평가자성명
                    oDS_PS_HR412L.SetValue("U_ColReg13", i, oRecordSet.Fields.Item("MSTCOD3").Value); //종합평가자사번
                    oDS_PS_HR412L.SetValue("U_ColReg14", i, oRecordSet.Fields.Item("FULLNAM3").Value); //종합평가자성명

                    oRecordSet.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// PS_HR412_etBaseForm
        /// </summary>
        private void PS_HR412_SetBaseForm(string oStatus)
        {
            string errMessage = string.Empty; 
            int i;
            string sQry;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            string Param06;
            string Param07;
            string Param08;
            string Param09;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                for (i = 1; i <= oMat01.RowCount; i++)
                {
                    if (oMat01.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
                    {
                        Param01 = oMat01.Columns.Item("MSTCOD").Cells.Item(i).Specific.Value;
                        Param02 = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(); //사업장
                        Param03 = oForm.Items.Item("Year").Specific.Value; //평가년도
                        Param04 = oForm.Items.Item("Number").Specific.Value.ToString().Trim(); //평가차수
                        Param05 = oForm.Items.Item("Evaluate").Specific.Value.ToString().Trim();  //평가권한
                        Param06 = oForm.Items.Item("Group").Specific.Value.ToString().Trim(); //반장, 사원구분
                        Param07 = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim(); //평가자사번
                        Param08 = oForm.Items.Item("FULLNAME").Specific.Value.ToString().Trim(); //평가자성명
                        if (oForm.Items.Item("Evaluate").Specific.Value == "1" && oForm.Items.Item("Group").Specific.Value == "2")
                        {
                            Param09 = oForm.Items.Item("CallName").Specific.Value; //직책명
                        }
                        else
                        {
                            Param09 = oForm.Items.Item("JigNm").Specific.Value; //직급명
                        }
                        sQry = "EXEC PS_HR412_02 '" + oStatus + "', '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "', '" + Param08 + "', '" + Param09 + "'";

                        oRecordSet.DoQuery(sQry);
                        PSH_Globals.SBO_Application.StatusBar.SetText("데이터 수정 완료");
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// PasswordChk
        /// </summary>
        /// <returns></returns>
        private bool PS_HR412_PasswordChk(SAPbouiCOM.ItemEvent pVal)
        {
            bool returnValue = false;
            string sQry;
            string MSTCOD;
            string PassWd;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                PassWd = oForm.Items.Item("PassWd").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(MSTCOD.ToString().Trim()))
                {
                    errMessage = "사번이 없습니다. 입력바랍니다.";
                    throw new Exception();
                }

                sQry = "Select Count(*) From Z_PS_HRPASS Where MSTCOD = '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
                sQry += " And  BPLId = '" + oForm.Items.Item("BPLId").Specific.Value + "' ";
                sQry += " And  PassWd = '" + oForm.Items.Item("PassWd").Specific.Value + "' ";
                RecordSet01.DoQuery(sQry);

                if (Convert.ToDouble(RecordSet01.Fields.Item(0).Value.ToString().Trim()) <= 0)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
            return returnValue;
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

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_HR412_PasswordChk(pVal) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("패스워드가 틀렸습니다. 확인바랍니다.");
                                oForm.Items.Item("PassWd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            else
                            {
                                PS_HR412_MTX01();
                            }
                        }
                    }

                    if (pVal.ItemUID == "Btn02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_HR412_PasswordChk(pVal) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("패스워드가 틀렸습니다. 확인바랍니다.");
                                oForm.Items.Item("PassWd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            else
                            {
                                PS_HR412_SetBaseForm("수정");
                                PS_HR412_MTX01();
                            }
                        }
                    }

                    if (pVal.ItemUID == "Btn03") //평가자 그룹핑 삭제
                    {
                        if (PS_HR412_PasswordChk(pVal) == false)
                        {
                            PSH_Globals.SBO_Application.MessageBox("패스워드가 틀렸습니다. 확인바랍니다.");
                            oForm.Items.Item("PassWd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            PS_HR412_SetBaseForm("삭제");
                            PS_HR412_MTX01();
                        }
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "Year")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "MSTCOD")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "TeamCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("TeamCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "RspCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("RspCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "ClsCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ClsCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
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
                if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02")
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
                            oMat01.SelectRow(pVal.Row, true, false); //메트릭스 한줄선택시 반전시켜주는 구문
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Year")
                        {
                            if (!string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                            {
                                sQry = "select U_Number from [@PS_HR410H] a";
                                sQry += " Where Isnull(a.U_OpenYN,'N') = 'Y' and isnull(a.U_CloseYN,'N') = 'N' ";
                                sQry += " and a.U_BPLId = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "' ";
                                sQry += " and a.U_Year = '" + oForm.Items.Item("Year").Specific.Value.ToString().Trim() + "' ";
                                oRecordSet.DoQuery(sQry);

                                if (!string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value.ToString().Trim()))
                                {
                                    oForm.Items.Item("Number").Specific.Select(oRecordSet.Fields.Item(0).Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                            }
                        }
                        else if (pVal.ItemUID == "MSTCOD")
                        {
                            sQry = "Select FULLNAME = t.U_FULLNAME, ";
                            sQry += " JigNm = Isnull((Select U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b on a.Code = b.Code  ";
                            sQry += " Where a.Name = '직급코드' And b.U_Code = t.U_JIGCOD ),''), ";
                            sQry += " CallName = Isnull((Select U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b on a.Code = b.Code  ";
                            sQry += " Where a.Name = '전문직호칭' And b.U_Code = t.U_CallName ),''), ";
                            sQry += " JicNm = Isnull((SELECT name FROM [OHPS]  ";
                            sQry += " Where posID = t.U_position ),'') ";
                            sQry += " From [@PH_PY001A] t Where Code =  '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "' ";
                            oRecordSet.DoQuery(sQry);

                            oForm.Items.Item("FULLNAME").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                            oForm.Items.Item("JigNm").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
                            oForm.Items.Item("CallName").Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();
                            oForm.Items.Item("JicNm").Specific.Value = oRecordSet.Fields.Item(3).Value.ToString().Trim();
                        }
                        else if (pVal.ItemUID == "TeamCode")
                        {
                            sQry = "Select b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code Where a.Name = '부서' and b.U_Code = '" + oForm.Items.Item("TeamCode").Specific.Value + "'";
                            oRecordSet.DoQuery(sQry);

                            oForm.Items.Item("TeamNm").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        }
                        else if (pVal.ItemUID == "RspCode")
                        {
                            sQry = "Select b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code Where a.Name = '담당' and b.U_Code = '" + oForm.Items.Item("RspCode").Specific.Value + "'";
                            oRecordSet.DoQuery(sQry);

                            oForm.Items.Item("RspNm").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        }
                        else if (pVal.ItemUID == "ClsCode")
                        {
                            sQry = "Select b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code Where a.Name = '반' and b.U_Code = '" + oForm.Items.Item("ClsCode").Specific.Value + "'";
                            oRecordSet.DoQuery(sQry);

                            oForm.Items.Item("ClsNm").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
                oForm.Update();
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
                    SubMain.Remove_Forms(oFormUniqueID);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_HR412L);
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
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당

            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "ItemCode")
                    {
                        if (oDataTable01 != null) //SelectedObjects 가 null이 아닐때만 실행(ChooseFromList 팝업창을 취소했을 때 미실행)
                        {
                            if (pVal.ItemUID == "ItemCode")
                            {
                                oForm.DataSources.UserDataSources.Item("ItemCode").Value = oDataTable01.Columns.Item("ItemCode").Cells.Item(0).Value;
                                oForm.DataSources.UserDataSources.Item("ItemName").Value = oDataTable01.Columns.Item("ItemName").Cells.Item(0).Value;
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
            }
        }

        /// <summary>
        /// Raise_EVENT_DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string Chk = null;

            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.Row == 0 && pVal.ColUID == "CHK")
                    {
                        oMat01.FlushToDataSource();
                        if (string.IsNullOrEmpty(oDS_PS_HR412L.GetValue("U_ColReg01", 0).ToString().Trim()) || oDS_PS_HR412L.GetValue("U_ColReg01", 0).ToString().Trim() == "N")
                        {
                            Chk = "Y";
                        }
                        else if (oDS_PS_HR412L.GetValue("U_ColReg01", 0).ToString().Trim() == "Y")
                        {
                            Chk = "N";
                        }
                        for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            oDS_PS_HR412L.SetValue("U_ColReg01", i, Chk);
                        }
                        oMat01.LoadFromDataSource();
                    }
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
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
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
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
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
                if (BusinessObjectInfo.BeforeAction == true)
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
                else if (BusinessObjectInfo.BeforeAction == false)
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
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
    }
}

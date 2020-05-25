using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// ���ڱݽ�û��
    /// </summary>
    internal class PH_PY305 : PSH_BaseClass
    {
        public string oFormUniqueID;
        public SAPbouiCOM.Matrix oMat01;

        //public SAPbouiCOM.Form oForm;

        //private SAPbouiCOM.DBDataSource oDS_PH_PY305A; //������
        private SAPbouiCOM.DBDataSource oDS_PH_PY305B; //��϶���

        private string oLastItemUID01; //Ŭ�������� ������ ������ ������ Uid��
        private string oLastColUID01; //�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
        private int oLastColRow01; //�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

        /// <summary>
        /// Form ȣ��
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY305.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY305_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY305");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy="DocEntry" '//UDO����϶�

                oForm.Freeze(true);
                PH_PY305_CreateItems();
                PH_PY305_EnableMenus();
                PH_PY305_SetDocument(oFromDocEntry01);

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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //�޸� ����
            }
        }

        /// <summary>
        /// ȭ�� Item ����
        /// </summary>
        private void PH_PY305_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                oDS_PH_PY305B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                // ��Ʈ���� ��ü �Ҵ�
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                // �����
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // �����ȣ
                oForm.DataSources.UserDataSources.Add("SCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SCode").Specific.DataBind.SetBound(true, "", "SCode");

                // �������
                oForm.DataSources.UserDataSources.Add("SName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("SName").Specific.DataBind.SetBound(true, "", "SName");

                // �⵵
                oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");
                oForm.Items.Item("StdYear").Specific.VALUE = DateTime.Now.ToString("yyyy");

                // �б�
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("", "");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("01", "1/4 Ȥ�� 1�б�");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("02", "2/4");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("03", "3/4 Ȥ�� 2�б�");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("04", "4/4");
                oForm.Items.Item("Quarter").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Quarter").DisplayDesc = true;

                // ȸ��
                oForm.Items.Item("Count").Specific.ValidValues.Add("%", "��ü");
                oForm.Items.Item("Count").Specific.ValidValues.Add("01", "1��");
                oForm.Items.Item("Count").Specific.ValidValues.Add("02", "2��");
                oForm.Items.Item("Count").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Count").DisplayDesc = true;

                // �Ҽ�
                oForm.DataSources.UserDataSources.Add("Team", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("Team").Specific.DataBind.SetBound(true, "", "Team");

                // ����
                oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

                // �Ի�����
                oForm.DataSources.UserDataSources.Add("startDat", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("startDat").Specific.DataBind.SetBound(true, "", "startDat");

                // ���бݰ�
                oForm.DataSources.UserDataSources.Add("FeeTot", SAPbouiCOM.BoDataType.dt_SUM, 50);
                oForm.Items.Item("FeeTot").Specific.DataBind.SetBound(true, "", "FeeTot");

                // ��ϱݰ�
                oForm.DataSources.UserDataSources.Add("TuiTot", SAPbouiCOM.BoDataType.dt_SUM, 50);
                oForm.Items.Item("TuiTot").Specific.DataBind.SetBound(true, "", "TuiTot");

                // �Ѱ�
                oForm.DataSources.UserDataSources.Add("Total", SAPbouiCOM.BoDataType.dt_SUM, 50);
                oForm.Items.Item("Total").Specific.DataBind.SetBound(true, "", "Total");


                // ��Ʈ���� SET
                // ����
                oMat01.Columns.Item("Sex").ValidValues.Add("", "");
                oMat01.Columns.Item("Sex").ValidValues.Add("01", "����");
                oMat01.Columns.Item("Sex").ValidValues.Add("02", "����");
                oMat01.Columns.Item("Sex").DisplayDesc = true;

                // �г�
                oMat01.Columns.Item("Grade").ValidValues.Add("", "");
                oMat01.Columns.Item("Grade").ValidValues.Add("01", "1�г�");
                oMat01.Columns.Item("Grade").ValidValues.Add("02", "2�г�");
                oMat01.Columns.Item("Grade").ValidValues.Add("03", "3�г�");
                oMat01.Columns.Item("Grade").ValidValues.Add("04", "4�г�");
                oMat01.Columns.Item("Grade").ValidValues.Add("05", "5�г�");
                oMat01.Columns.Item("Grade").DisplayDesc = true;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY305_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //�޸� ����
            }
        }

        /// <summary>
        /// �޴� ������ Enable
        /// </summary>
        private void PH_PY305_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false); // ����
                oForm.EnableMenu("1286", false); // �ݱ�
                oForm.EnableMenu("1287", false); // ����
                oForm.EnableMenu("1285", false); // ����
                oForm.EnableMenu("1284", false); // ���
                oForm.EnableMenu("1293", false); // �����
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY305_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// ȭ�鼼��
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY305_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY305_FormItemEnabled();
                    //Call PH_PY305_AddMatrixRow(0, True) '//UDO����϶�
                }
                else
                {
                    //oForm.Mode = fm_FIND_MODE
                    //PH_PY305_FormItemEnabled
                    //oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
                    //oForm.Items("1").Click ct_Regular
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY305_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// ȭ���� ������ Enable ����
        /// </summary>
        private void PH_PY305_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY305_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ������ ��ȸ
        /// </summary>
        private void PH_PY305_MTX01()
        {
            int i = 0;
            string sQry = null;
            short ErrNum = 0;
            double FeeTot = 0;    // ���бݰ�
            double TuiTot = 0;    // ��ϱݰ�
            double Total = 0;     // �Ѱ�
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string CLTCOD = string.Empty; //�����
            string SCode = string.Empty;
            string StdYear = string.Empty;
            string Quarter = string.Empty;
            string Count = string.Empty;

            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet.RecordCount, false);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();   // �����
                SCode = oForm.Items.Item("SCode").Specific.VALUE.Trim();
                StdYear = oForm.Items.Item("StdYear").Specific.VALUE.Trim();
                Quarter = oForm.Items.Item("Quarter").Specific.VALUE.Trim();
                Count = oForm.Items.Item("Count").Specific.VALUE.Trim();

                sQry = "            EXEC [PH_PY305_01] ";
                sQry = sQry + "'" + CLTCOD + "',"; //�����
                sQry = sQry + "'" + SCode + "',";
                sQry = sQry + "'" + StdYear + "',";
                sQry = sQry + "'" + Quarter + "',";
                sQry = sQry + "'" + Count + "'";

                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY305B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                oForm.Items.Item("Team").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;        // ��û��-�Ҽ�
                oForm.Items.Item("CntcName").Specific.VALUE = oRecordSet.Fields.Item("CntcName").Value;    // ��û��-����
                oForm.Items.Item("startDat").Specific.VALUE = oRecordSet.Fields.Item("startDat").Value;    // ��û��-�Ի�����

                if (oRecordSet.RecordCount == 0)
                {
                    ErrNum = 1;
                    oMat01.Clear();
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY305B.Size)
                    {
                        oDS_PH_PY305B.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY305B.Offset = i;

                    oDS_PH_PY305B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY305B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Name").Value);
                    oDS_PH_PY305B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("BirthDat").Value);
                    oDS_PH_PY305B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("Sex").Value);
                    oDS_PH_PY305B.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("SchName").Value);
                    oDS_PH_PY305B.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("Grade").Value);
                    oDS_PH_PY305B.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("EntFee").Value);
                    oDS_PH_PY305B.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("Tuition").Value);

                    FeeTot = FeeTot + oRecordSet.Fields.Item("EntFee").Value;
                    TuiTot = TuiTot + oRecordSet.Fields.Item("Tuition").Value;

                    oRecordSet.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet.RecordCount + "�� ��ȸ��...!";

                }
                Total = FeeTot + TuiTot;

                oForm.Items.Item("FeeTot").Specific.VALUE = FeeTot;
                oForm.Items.Item("TuiTot").Specific.VALUE = TuiTot;
                oForm.Items.Item("Total").Specific.VALUE = Total;

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                ProgBar01.Stop();

            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("��ȸ ����� �����ϴ�. Ȯ���ϼ���.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY305_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// FlushToItemValue(������� Event�� ���� ȭ�� Item�� �������� ����)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PH_PY305_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                switch (oUID)
                {
                    case "SCode":
                        oForm.Items.Item("SName").Specific.VALUE = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("SCode").Specific.VALUE + "'", "");
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY305_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// ����Ʈ ���
        /// </summary>
        private void PH_PY305_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty; //�����
            string SCode = string.Empty;
            string StdYear = string.Empty;
            string Quarter = string.Empty;
            string Count = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();   // �����
            SCode = oForm.Items.Item("SCode").Specific.VALUE.Trim();
            StdYear = oForm.Items.Item("StdYear").Specific.VALUE.Trim();
            Quarter = oForm.Items.Item("Quarter").Specific.VALUE.Trim();
            Count = oForm.Items.Item("Count").Specific.VALUE.Trim();

            try
            {
                WinTitle = "[PH_PY305] ���ڱݽ�û����(���κ�)";
                ReportName = "PH_PY305_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();//Parameter List
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                // Formula
                //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'")));

                // Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //�����
                dataPackParameter.Add(new PSH_DataPackClass("@SCode", SCode));
                dataPackParameter.Add(new PSH_DataPackClass("@StdYear", StdYear));
                dataPackParameter.Add(new PSH_DataPackClass("@Quarter", Quarter));
                dataPackParameter.Add(new PSH_DataPackClass("@Count", Count));

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                                                             //    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
        /// ITEM_PRESSED �̺�Ʈ
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent ��ü</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "PH_PY305")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch") //��ȸ ��ư
                    {
                        if (PH_PY305_DataValidCheck() == true)
                        {
                            PH_PY305_MTX01();
                        }
                    }
                    else if (pVal.ItemUID == "BtnPrint")
                    {
                        if (PH_PY305_DataValidCheck() == true)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PH_PY305_Print_Report01);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PH_PY305")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
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
        /// KEY_DOWN �̺�Ʈ
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent ��ü</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD", ""); //���
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ShiftDatCd", ""); //�ٹ�����
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "GNMUJOCd", ""); //�ٹ���
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// GOT_FOCUS �̺�Ʈ
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent ��ü</param>
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY305_DataValidCheck()
        {
            bool functionReturnValue = false;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.VALUE))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("������� �ʼ��Դϴ�.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                // �⵵
                if (string.IsNullOrEmpty(oForm.Items.Item("StdYear").Specific.VALUE.Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("�⵵�� �ʼ��Դϴ�.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("StdYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                // �б�
                if (string.IsNullOrEmpty(oForm.Items.Item("Quarter").Specific.VALUE.Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("�б�� �ʼ��Դϴ�.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Quarter").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY305_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                functionReturnValue = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// COMBO_SELECT �̺�Ʈ
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent ��ü</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PH_PY305_FlushToItemValue(pVal.ItemUID, 0, "");
                }
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
        /// CLICK �̺�Ʈ
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent ��ü</param>
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DOUBLE_CLICK �̺�Ʈ
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent ��ü</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_DOUBLE_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// MATRIX_LINK_PRESSED �̺�Ʈ
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent ��ü</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LINK_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// VALIDATE �̺�Ʈ
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent ��ü</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                        }
                        else
                        {
                            PH_PY305_FlushToItemValue(pVal.ItemUID, 0, "");

                            //if (pVal.ItemUID == "MSTCOD")
                            //{
                            //    oForm.Items.Item("MSTNAM").Specific.VALUE = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.VALUE + "'", ""); //����
                            //}
                            //else if (pVal.ItemUID == "ShiftDatCd")
                            //{
                            //    oForm.Items.Item("ShiftDatNm").Specific.VALUE = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L] AS T0", "'" + oForm.Items.Item("ShiftDatCd").Specific.VALUE + "'", " AND T0.Code = 'P154' AND T0.U_UseYN = 'Y'"); //�ٹ�����

                            //}
                            //else if (pVal.ItemUID == "GNMUJOCd")
                            //{
                            //    oForm.Items.Item("GNMUJONm").Specific.VALUE = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L] AS T0", "'" + oForm.Items.Item("GNMUJOCd").Specific.VALUE + "'", " AND T0.Code = 'P155' AND T0.U_UseYN = 'Y'"); //�ٹ���
                            //}
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// MATRIX_LOAD �̺�Ʈ
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent ��ü</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PH_PY305_FormItemEnabled();
                    //PH_PY305_AddMatrixRow(oMat01.VisualRowCount);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// FORM_UNLOAD �̺�Ʈ
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent ��ü</param>
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
        /// CHOOSE_FROM_LIST �̺�Ʈ
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent ��ü</param>
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
                    //���� �ҽ�(VB6.0 �ּ�ó���Ǿ� ����)
                    //If (pval.ItemUID = "ItemCode") Then
                    //  Dim oDataTable01 As SAPbouiCOM.DataTable
                    //  Set oDataTable01 = pval.SelectedObjects
                    //  oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
                    //  Set oDataTable01 = Nothing
                    //End If
                    //If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
                    //  Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY305A", "U_CardCode,U_CardName")
                    //End If
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CHOOSE_FROM_LIST_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //���
                            break;
                        case "1286": //�ݱ�
                            break;
                        case "1293": //�����
                            break;
                        case "1281": //ã��
                            break;
                        case "1282": //�߰�
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //���ڵ��̵���ư
                            break;
                        case "7169": //���� ��������
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //���
                            break;
                        case "1286": //�ݱ�
                            break;
                        case "1293": //�����
                            //Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //ã��
                            break;
                        case "1282": //�߰�
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                        case "7169": //���� ��������
                            //���� �������� ���� ó��_S
                            oForm.Freeze(true);
                            oDS_PH_PY305B.RemoveRecord(oDS_PH_PY305B.Size - 1);
                            oMat01.LoadFromDataSource();
                            oForm.Freeze(false);
                            //���� �������� ���� ó��_E
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

        /// <summary>
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            //string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        #region Raise_FormMenuEvent (�����׽�Ʈ �� �ּ� ���� �ʿ�, 2019.05.17 �۸��)
        //		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			string sQry = null;
        //			SAPbobsCOM.Recordset RecordSet01 = null;
        //			RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			////BeforeAction = True
        //			if ((pval.BeforeAction == true)) {

        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {

        //			}
        //			return;
        //			Raise_FormMenuEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormDataEvent (�����׽�Ʈ �� �ּ� ���� �ʿ�, 2019.05.17 �۸��)
        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////BeforeAction = True
        //			if ((BusinessObjectInfo.BeforeAction == true)) {

        //			////BeforeAction = False
        //			} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //			}
        //			return;
        //			Raise_FormDataEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_RightClickEvent (�����׽�Ʈ �� �ּ� ���� �ʿ�, 2019.05.17 �۸��)
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {
        //			}

        //			return;
        //			Raise_RightClickEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormItemEvent (�����׽�Ʈ �� �ּ� ���� �ʿ�, 2019.05.17 �۸��)
        //		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_ITEM_PRESSED_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_KEY_DOWN_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {


        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_CLICK_Error:

        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_COMBO_SELECT_Error:
        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_DOUBLE_CLICK_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oForm.Freeze(true);

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			oForm.Freeze(false);

        //			return;
        //			Raise_EVENT_VALIDATE_Error:

        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LOAD_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_RESIZE_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}


        //		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			return;
        //			Raise_EVENT_GOT_FOCUS_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_FORM_UNLOAD_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion
    }
}




//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
// // ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_HR_Addon
//{
//	internal class PH_PY305
//	{
//////********************************************************************************
//////  File           : PH_PY305.cls
//////  Module         : �λ���� > ��Ÿ
//////  Desc           : ���ڱݽ�û��
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Matrix oMat1;

//		private SAPbouiCOM.DBDataSource oDS_PH_PY305A;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY305B;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY305.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY305_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY305");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			//    oForm.DataBrowser.BrowseBy = "DocEntry"

//			oForm.Freeze(true);
//			PH_PY305_CreateItems();
//			PH_PY305_EnableMenus();
//			PH_PY305_SetDocument(oFromDocEntry01);
//			//    Call PH_PY305_FormResize

//			oForm.Update();
//			oForm.Freeze(false);

//			oForm.Visible = true;
//			//UPGRADE_NOTE: oXmlDoc ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			return;
//			LoadForm_Error:

//			oForm.Update();
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oXmlDoc ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			//UPGRADE_NOTE: oForm ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oForm = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private bool PH_PY305_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;

//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//    Set oDS_PH_PY305A = oForm.DataSources.DBDataSources("@PH_PY305A")
//			oDS_PH_PY305B = oForm.DataSources.DBDataSources("@PS_USERDS01");

//			oMat1 = oForm.Items.Item("Mat01").Specific;

//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
//			oMat1.AutoResizeColumns();

//			//----------��ȸ ����----------
//			//�����_S
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
//			//�����_E

//			//�����ȣ_S
//			oForm.DataSources.UserDataSources.Add("SCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("SCode").Specific.DataBind.SetBound(true, "", "SCode");
//			//�����ȣ_E

//			//�������_S
//			oForm.DataSources.UserDataSources.Add("SName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("SName").Specific.DataBind.SetBound(true, "", "SName");
//			//�������_E

//			//�⵵_S
//			oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");
//			//�⵵_E

//			//�б�_S
//			oForm.DataSources.UserDataSources.Add("Quarter", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Quarter").Specific.DataBind.SetBound(true, "", "Quarter");
//			//�б�_E
//			//----------��ȸ ����----------

//			//�Ҽ�_S
//			oForm.DataSources.UserDataSources.Add("Team", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Team").Specific.DataBind.SetBound(true, "", "Team");
//			//�Ҽ�_E

//			//����_S
//			oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");
//			//����_E

//			//�Ի�����_S
//			oForm.DataSources.UserDataSources.Add("startDat", SAPbouiCOM.BoDataType.dt_DATE);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("startDat").Specific.DataBind.SetBound(true, "", "startDat");
//			//�Ի�����_E

//			//���бݰ�_S
//			oForm.DataSources.UserDataSources.Add("FeeTot", SAPbouiCOM.BoDataType.dt_SUM, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("FeeTot").Specific.DataBind.SetBound(true, "", "FeeTot");
//			//���бݰ�_E

//			//��ϱݰ�_S
//			oForm.DataSources.UserDataSources.Add("TuiTot", SAPbouiCOM.BoDataType.dt_SUM, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("TuiTot").Specific.DataBind.SetBound(true, "", "TuiTot");
//			//��ϱݰ�_E

//			//�Ѱ�_S
//			oForm.DataSources.UserDataSources.Add("Total", SAPbouiCOM.BoDataType.dt_SUM, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Total").Specific.DataBind.SetBound(true, "", "Total");
//			//�Ѱ�_E

//			////----------------------------------------------------------------------------------------------
//			//// �⺻����
//			////----------------------------------------------------------------------------------------------

//			//�����
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			//�б�
//			oCombo = oForm.Items.Item("Quarter").Specific;
//			oCombo.ValidValues.Add("", "");
//			oCombo.ValidValues.Add("01", "1/4 Ȥ�� 1�б�");
//			oCombo.ValidValues.Add("02", "2/4");
//			oCombo.ValidValues.Add("03", "3/4 Ȥ�� 2�б�");
//			oCombo.ValidValues.Add("04", "4/4");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("Quarter").DisplayDesc = true;

//			//��Ʈ����-����
//			oColumn = oMat1.Columns.Item("Sex");
//			oColumn.ValidValues.Add("", "");
//			oColumn.ValidValues.Add("01", "����");
//			oColumn.ValidValues.Add("02", "����");
//			oColumn.DisplayDesc = true;

//			//��Ʈ����-�г�
//			oColumn = oMat1.Columns.Item("Grade");
//			oColumn.ValidValues.Add("", "");
//			oColumn.ValidValues.Add("01", "1�г�");
//			oColumn.ValidValues.Add("02", "2�г�");
//			oColumn.ValidValues.Add("03", "3�г�");
//			oColumn.ValidValues.Add("04", "4�г�");
//			oColumn.DisplayDesc = true;


//			//UPGRADE_NOTE: oEdit ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY305_CreateItems_Error:

//			//UPGRADE_NOTE: oEdit ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY305_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY305_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			//    Call oForm.EnableMenu("1283", False)         '// ����
//			//    Call oForm.EnableMenu("1287", True)          '// ����
//			//'    Call oForm.EnableMenu("1286", True)         '// �ݱ�
//			//    Call oForm.EnableMenu("1284", True)         '// ���
//			//    Call oForm.EnableMenu("1293", True)         '// �����

//			return;
//			PH_PY305_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY305_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY305_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY305_FormItemEnabled();
//				//        Call PH_PY305_AddMatrixRow
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY305_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY305_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY305_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY305_FormItemEnabled()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbouiCOM.ComboBox oCombo = null;
//			string CLTCOD = null;

//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				//�� DocEntry ����
//				//        Call PH_PY305_FormClear

//				//// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				//�⵵ ����
//				//        Call oDS_PH_PY305A.setValue("U_StdYear", 0, Format(Date, "YYYY"))
//				//UPGRADE_WARNING: oForm.Items(StdYear).Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("StdYear").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYY");

//				oForm.EnableMenu("1281", true);
//				////����ã��
//				oForm.EnableMenu("1282", false);
//				////�����߰�

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				//// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				oForm.EnableMenu("1281", false);
//				////����ã��
//				oForm.EnableMenu("1282", true);
//				////�����߰�

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				//// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//				oForm.EnableMenu("1281", true);
//				////����ã��
//				oForm.EnableMenu("1282", true);
//				////�����߰�

//			}
//			oForm.Freeze(false);
//			return;
//			PH_PY305_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY305_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			short loopCount = 0;
//			//For Loop �� (VALIDATE Event���� ���)
//			string GovID1 = null;
//			//�ֹε�Ϲ�ȣ ���ڸ�(VALIDATE Event���� ���)
//			string GovID2 = null;
//			//�ֹε�Ϲ�ȣ ���ڸ�(VALIDATE Event���� ���)
//			string GovID = null;
//			//�ֹε�Ϲ�ȣ ��ü(VALIDATE Event���� ���)
//			string Sex = null;
//			//����(VALIDATE Event���� ���)
//			short PayCnt = 0;
//			//����Ƚ��(VALIDATE Event���� ���)
//			double FeeTot = 0;
//			//���бݰ�(VALIDATE Event���� ���)
//			double TuiTot = 0;
//			//��ϱݰ�(VALIDATE Event���� ���)
//			double Total = 0;
//			//�Ѱ�(VALIDATE Event���� ���)

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1

//					if (pval.BeforeAction == true) {
//						//                If pval.ItemUID = "1" Then
//						//                    If oForm.Mode = fm_ADD_MODE Then
//						//                        If PH_PY305_DataValidCheck = False Then
//						//                            BubbleEvent = False
//						//                        End If
//						//
//						//                        '//�ؾ����� �۾�
//						//                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
//						//                        If PH_PY305_DataValidCheck = False Then
//						//                            BubbleEvent = False
//						//                        End If
//						//                        '//�ؾ����� �۾�
//						//
//						//                    ElseIf oForm.Mode = fm_OK_MODE Then
//						//                    End If
//						//                End If
//						if (pval.ItemUID == "BtnSearch") {

//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (PH_PY305_DataValidCheck() == false) {
//									BubbleEvent = false;
//									return;
//								}
//								//
//								//                        '//�ؾ����� �۾�
//								PH_PY305_MTX01();

//							}

//						}
//						if (pval.ItemUID == "BtnPrint") {

//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (PH_PY305_DataValidCheck() == false) {
//									BubbleEvent = false;
//									return;
//								}
//								//
//								//                        '//�ؾ����� �۾�
//								PH_PY305_Print_Report01();

//							}

//						}

//						//
//					} else if (pval.BeforeAction == false) {
//						//                If pval.ItemUID = "1" Then
//						//                    If oForm.Mode = fm_ADD_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY305_FormItemEnabled
//						//                            Call PH_PY305_AddMatrixRow
//						//                        End If
//						//                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY305_FormItemEnabled
//						//                            Call PH_PY305_AddMatrixRow
//						//                        End If
//						//                    ElseIf oForm.Mode = fm_OK_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY305_FormItemEnabled
//						//                        End If
//						//                    End If
//						//                End If
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2

//					if (pval.BeforeAction == true) {

//						if (pval.ItemUID == "Mat01") {

//							if (pval.ColUID == "Name" & pval.CharPressed == Convert.ToDouble("9")) {

//								//UPGRADE_WARNING: oMat1.Columns.Item(Name).Cells(pval.Row).Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (string.IsNullOrEmpty(oMat1.Columns.Item("Name").Cells.Item(pval.Row).Specific.VALUE)) {
//									MDC_Globals.Sbo_Application.ActivateMenuItem("7425");
//									BubbleEvent = false;
//								}

//							}

//						} else if (pval.ItemUID == "SCode" & pval.CharPressed == Convert.ToDouble("9")) {

//							//UPGRADE_WARNING: oForm.Items(SCode).Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (string.IsNullOrEmpty(oForm.Items.Item("SCode").Specific.VALUE)) {
//								MDC_Globals.Sbo_Application.ActivateMenuItem("7425");
//								BubbleEvent = false;
//							}

//						}

//					} else if (pval.Before_Action == false) {

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat01":
//							if (pval.Row > 0) {
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = pval.ColUID;
//								oLastColRow = pval.Row;
//							}
//							break;
//						default:
//							oLastItemUID = pval.ItemUID;
//							oLastColUID = "";
//							oLastColRow = 0;
//							break;
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//					////4
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					////5
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {
//							//                    Call PH_PY305_AddMatrixRow
//							oMat1.AutoResizeColumns();
//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					if (pval.BeforeAction == true) {
//						switch (pval.ItemUID) {
//							case "Mat01":
//								if (pval.Row > 0) {
//									oMat1.SelectRow(pval.Row, true, false);
//								}
//								break;
//						}

//						switch (pval.ItemUID) {
//							case "Mat01":
//								if (pval.Row > 0) {
//									oLastItemUID = pval.ItemUID;
//									oLastColUID = pval.ColUID;
//									oLastColRow = pval.Row;
//								}
//								break;
//							default:
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = "";
//								oLastColRow = 0;
//								break;
//						}
//					} else if (pval.BeforeAction == false) {

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//					////8
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//					////9
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					////10
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {

//						if (pval.ItemChanged == true) {

//						}

//					} else if (pval.BeforeAction == false) {

//						if (pval.ItemChanged == true) {

//							switch (pval.ItemUID) {

//								case "SCode":
//									//Call oDS_PH_PY305A.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.Value & "'"))

//									//UPGRADE_WARNING: oForm.Items(SName).Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: MDC_GetData.Get_ReData() ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("SName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pval.ItemUID).Specific.VALUE + "'");
//									break;

//							}

//						}

//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						oMat1.LoadFromDataSource();

//						PH_PY305_FormItemEnabled();
//						PH_PY305_AddMatrixRow();
//						oMat1.AutoResizeColumns();

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//					////12
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//					////16
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					////17
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oDS_PH_PY305A ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY305A = null;
//						//UPGRADE_NOTE: oDS_PH_PY305B ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY305B = null;

//						//UPGRADE_NOTE: oMat1 ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//					////18
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//					////19
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//					////20
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//					////21
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {

//						oMat1.AutoResizeColumns();

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//					////22
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//					////23
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//					////27
//					if (pval.BeforeAction == true) {

//					} else if (pval.Before_Action == false) {
//						//                If pval.ItemUID = "Code" Then
//						//                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY305A", "Code")
//						//                End If
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//					////37
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//					////38
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_Drag:
//					////39
//					break;

//			}

//			//UPGRADE_NOTE: oCombo ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			return;
//			Raise_FormItemEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			oForm.Freeze((false));
//			//UPGRADE_NOTE: oCombo ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			int i = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short loopCount = 0;
//			double FeeTot = 0;
//			double TuiTot = 0;
//			double Total = 0;

//			oForm.Freeze(true);

//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						if (MDC_Globals.Sbo_Application.MessageBox("���� ȭ�鳻����ü�� ���� �Ͻðڽ��ϱ�? ������ �� �����ϴ�.", 2, "Yes", "No") == 2) {
//							BubbleEvent = false;
//							return;
//						}
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					case "1293":
//						break;
//					case "1281":
//						break;
//					case "1282":
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						break;
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY305_FormItemEnabled();
//						PH_PY305_AddMatrixRow();
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////����ã��
//						PH_PY305_FormItemEnabled();
//						PH_PY305_AddMatrixRow();
//						oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////�����߰�
//						PH_PY305_FormItemEnabled();
//						PH_PY305_AddMatrixRow();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY305_FormItemEnabled();
//						break;
//					case "1293":
//						//// �����

//						if (oMat1.RowCount != oMat1.VisualRowCount) {
//							oMat1.FlushToDataSource();

//							while ((i <= oDS_PH_PY305B.Size - 1)) {
//								if (string.IsNullOrEmpty(oDS_PH_PY305B.GetValue("U_LineNum", i))) {
//									oDS_PH_PY305B.RemoveRecord((i));
//									i = 0;
//								} else {
//									i = i + 1;
//								}
//							}

//							for (i = 0; i <= oDS_PH_PY305B.Size; i++) {
//								oDS_PH_PY305B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//							}

//							oMat1.LoadFromDataSource();
//						}
//						PH_PY305_AddMatrixRow();
//						break;

//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


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


//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//			}
//			switch (pval.ItemUID) {
//				case "Mat01":
//					if (pval.Row > 0) {
//						oLastItemUID = pval.ItemUID;
//						oLastColUID = pval.ColUID;
//						oLastColRow = pval.Row;
//					}
//					break;
//				default:
//					oLastItemUID = pval.ItemUID;
//					oLastColUID = "";
//					oLastColRow = 0;
//					break;
//			}
//			return;
//			Raise_RightClickEvent_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY305_AddMatrixRow()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			////[Mat1]
//			oMat1.FlushToDataSource();
//			oRow = oMat1.VisualRowCount;

//			if (oMat1.VisualRowCount > 0) {
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY305B.GetValue("U_Name", oRow - 1)))) {
//					if (oDS_PH_PY305B.Size <= oMat1.VisualRowCount) {
//						oDS_PH_PY305B.InsertRecord((oRow));
//					}
//					oDS_PH_PY305B.Offset = oRow;
//					oDS_PH_PY305B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//					oDS_PH_PY305B.SetValue("U_ColReg01", oRow, "");
//					oDS_PH_PY305B.SetValue("U_ColReg02", oRow, "");
//					oDS_PH_PY305B.SetValue("U_ColReg03", oRow, "");
//					oDS_PH_PY305B.SetValue("U_ColReg04", oRow, "");
//					oDS_PH_PY305B.SetValue("U_ColReg05", oRow, "");
//					oDS_PH_PY305B.SetValue("U_ColSum01", oRow, "");
//					oDS_PH_PY305B.SetValue("U_ColSum02", oRow, "");
//					oMat1.LoadFromDataSource();
//				} else {
//					oDS_PH_PY305B.Offset = oRow - 1;
//					oDS_PH_PY305B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//					oDS_PH_PY305B.SetValue("U_ColReg01", oRow - 1, "");
//					oDS_PH_PY305B.SetValue("U_ColReg02", oRow - 1, "");
//					oDS_PH_PY305B.SetValue("U_ColReg03", oRow - 1, "");
//					oDS_PH_PY305B.SetValue("U_ColReg04", oRow - 1, "");
//					oDS_PH_PY305B.SetValue("U_ColReg05", oRow - 1, "");
//					oDS_PH_PY305B.SetValue("U_ColSum01", oRow - 1, "");
//					oDS_PH_PY305B.SetValue("U_ColSum02", oRow - 1, "");
//					oMat1.LoadFromDataSource();
//				}
//			} else if (oMat1.VisualRowCount == 0) {
//				oDS_PH_PY305B.Offset = oRow;
//				oDS_PH_PY305B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//				oDS_PH_PY305B.SetValue("U_ColReg01", oRow, "");
//				oDS_PH_PY305B.SetValue("U_ColReg02", oRow, "");
//				oDS_PH_PY305B.SetValue("U_ColReg03", oRow, "");
//				oDS_PH_PY305B.SetValue("U_ColReg04", oRow, "");
//				oDS_PH_PY305B.SetValue("U_ColReg05", oRow, "");
//				oDS_PH_PY305B.SetValue("U_ColSum01", oRow, "");
//				oDS_PH_PY305B.SetValue("U_ColSum02", oRow, "");
//				oMat1.LoadFromDataSource();
//			}

//			oForm.Freeze(false);
//			return;
//			PH_PY305_AddMatrixRow_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY305_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY305_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY305'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY305_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY305_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY305_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//�����
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("������� �ʼ��Դϴ�.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//    '�����ȣ
//			//    If Trim(oForm.Items("SCode").Specific.VALUE) = "" Then
//			//        Sbo_Application.SetStatusBarMessage "�����ȣ�� �ʼ��Դϴ�.", bmt_Short, True
//			//        oForm.Items("SCode").CLICK ct_Regular
//			//        PH_PY305_DataValidCheck = False
//			//        Exit Function
//			//    End If

//			//�⵵
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("StdYear").Specific.VALUE))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("�⵵�� �ʼ��Դϴ�.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("StdYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//�б�
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("Quarter").Specific.VALUE))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("�б�� �ʼ��Դϴ�.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("Quarter").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}


//			//    '����
//			//    If oMat1.VisualRowCount > 1 Then
//			//        For i = 1 To oMat1.VisualRowCount - 1
//			//
//			//            '�б�
//			//            If oMat1.Columns("SchCls").Cells(i).Specific.Value = "" Then
//			//                Sbo_Application.SetStatusBarMessage "�б��� �ʼ��Դϴ�.", bmt_Short, True
//			//                oMat1.Columns("SchCls").Cells(i).CLICK ct_Regular
//			//                PH_PY305_DataValidCheck = False
//			//                Exit Function
//			//            End If
//			//
//			//            '�б���
//			//            If oMat1.Columns("SchName").Cells(i).Specific.Value = "" Then
//			//                Sbo_Application.SetStatusBarMessage "�б����� �ʼ��Դϴ�.", bmt_Short, True
//			//                oMat1.Columns("SchName").Cells(i).CLICK ct_Regular
//			//                PH_PY305_DataValidCheck = False
//			//                Exit Function
//			//            End If
//			//
//			//            '�г�
//			//            If oMat1.Columns("Grade").Cells(i).Specific.Value = "" Then
//			//                Sbo_Application.SetStatusBarMessage "�г��� �ʼ��Դϴ�.", bmt_Short, True
//			//                oMat1.Columns("Grade").Cells(i).CLICK ct_Regular
//			//                PH_PY305_DataValidCheck = False
//			//                Exit Function
//			//            End If
//			//
//			//            'ȸ��
//			//            If oMat1.Columns("Count").Cells(i).Specific.Value = "" Then
//			//                Sbo_Application.SetStatusBarMessage "ȸ���� �ʼ��Դϴ�.", bmt_Short, True
//			//                oMat1.Columns("Count").Cells(i).CLICK ct_Regular
//			//                PH_PY305_DataValidCheck = False
//			//                Exit Function
//			//            End If
//			//
//			//        Next
//			//    Else
//			//        Sbo_Application.SetStatusBarMessage "���� �����Ͱ� �����ϴ�.", bmt_Short, True
//			//        PH_PY305_DataValidCheck = False
//			//        Exit Function
//			//    End If

//			oMat1.FlushToDataSource();
//			//// Matrix ������ �� ����(DB �����)
//			if (oDS_PH_PY305B.Size > 1)
//				oDS_PH_PY305B.RemoveRecord((oDS_PH_PY305B.Size - 1));

//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;
//			return functionReturnValue;


//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			PH_PY305_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY305_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY305_MTX01()
//		{

//			////��Ʈ������ ������ �ε�

//			int i = 0;
//			string sQry = null;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;
//			string Param04 = null;

//			double FeeTot = 0;
//			//���бݰ�
//			double TuiTot = 0;
//			//��ϱݰ�
//			double Total = 0;
//			//�Ѱ�

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oForm.Items.Item("SCode").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = oForm.Items.Item("StdYear").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param04 = oForm.Items.Item("Quarter").Specific.VALUE;

//			SAPbouiCOM.ProgressBar ProgressBar01 = null;
//			ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet.RecordCount, false);

//			sQry = "EXEC PH_PY305_01 '" + Param01 + "','" + Param02 + "','" + Param03 + "','" + Param04 + "'";
//			oRecordSet.DoQuery(sQry);

//			//UPGRADE_WARNING: oForm.Items(Team).Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Team").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
//			//��û��-�Ҽ�
//			//UPGRADE_WARNING: oForm.Items(CntcName).Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("CntcName").Specific.VALUE = oRecordSet.Fields.Item("CntcName").Value;
//			//��û��-����
//			//UPGRADE_WARNING: oForm.Items(startDat).Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("startDat").Specific.VALUE = oRecordSet.Fields.Item("startDat").Value;
//			//��û��-�Ի�����

//			oMat1.Clear();
//			oMat1.FlushToDataSource();
//			oMat1.LoadFromDataSource();

//			if ((oRecordSet.RecordCount == 0)) {
//				oMat1.Clear();
//				goto PH_PY305_MTX01_Exit;
//			}

//			for (i = 0; i <= oRecordSet.RecordCount - 1; i++) {
//				if (i != 0) {
//					oDS_PH_PY305B.InsertRecord((i));
//				}
//				oDS_PH_PY305B.Offset = i;
//				oDS_PH_PY305B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//				oDS_PH_PY305B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Name").Value);
//				oDS_PH_PY305B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("BirthDat").Value);
//				oDS_PH_PY305B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("Sex").Value);
//				oDS_PH_PY305B.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("SchName").Value);
//				oDS_PH_PY305B.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("Grade").Value);
//				oDS_PH_PY305B.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("EntFee").Value);
//				oDS_PH_PY305B.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("Tuition").Value);

//				FeeTot = FeeTot + oRecordSet.Fields.Item("EntFee").Value;
//				TuiTot = TuiTot + oRecordSet.Fields.Item("Tuition").Value;

//				oRecordSet.MoveNext();
//				ProgressBar01.Value = ProgressBar01.Value + 1;
//				ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "�� ��ȸ��...!";

//			}

//			Total = FeeTot + TuiTot;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("FeeTot").Specific.VALUE = FeeTot;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("TuiTot").Specific.VALUE = TuiTot;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Total").Specific.VALUE = Total;

//			oMat1.LoadFromDataSource();
//			oMat1.AutoResizeColumns();
//			oForm.Update();

//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY305_MTX01_Exit:
//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			if ((ProgressBar01 != null)) {
//				ProgressBar01.Stop();
//			}
//			MDC_Com.MDC_GF_Message(ref "����� �������� �ʽ��ϴ�.", ref "W");
//			return;
//			PH_PY305_MTX01_Error:
//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY305_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY305_Validate(string ValidateType)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = true;
//			object i = null;
//			int j = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY305A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY305A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("�ش繮���� �ٸ�����ڿ� ���� ��ҵǾ����ϴ�. �۾��� �����Ҽ� �����ϴ�.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY305_Validate_Exit;
//			}
//			//
//			if (ValidateType == "����") {

//			} else if (ValidateType == "�����") {

//			} else if (ValidateType == "���") {

//			}
//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY305_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY305_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY305_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY305_Print_Report01()
//		{

//			string DocNum = null;
//			short ErrNum = 0;
//			string WinTitle = null;
//			string ReportName = null;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			string CLTCOD = null;
//			string sCode = null;
//			string StdYear = null;
//			string Quarter = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


//			/// ODBC ���� üũ
//			if (ConnectODBC() == false) {
//				goto PH_PY305_Print_Report01_Error;
//			}


//			////���� MOVE , Trim ��Ű��..
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sCode = oForm.Items.Item("SCode").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			StdYear = oForm.Items.Item("StdYear").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Quarter = oForm.Items.Item("Quarter").Specific.VALUE;


//			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

//			WinTitle = "[PH_PY305] ���ڱݽ�û��";
//			ReportName = "PH_PY305_01.rpt";
//			MDC_Globals.gRpt_Formula = new string[2];
//			MDC_Globals.gRpt_Formula_Value = new string[2];
//			MDC_Globals.gRpt_SRptSqry = new string[2];
//			MDC_Globals.gRpt_SRptName = new string[2];
//			MDC_Globals.gRpt_SFormula = new string[2, 2];
//			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

//			/// Formula �����ʵ�

//			//    gRpt_Formula(1) = "CLTCOD"
//			//    sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & CLTCOD & "'"
//			//    Call oRecordSet.DoQuery(sQry)
//			//    gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
//			//
//			//    gRpt_Formula(2) = "YM"
//			//    gRpt_Formula_Value(2) = Format(YM, "####��##��")
//			//
//			//    gRpt_Formula(3) = "DocDate"
//			//    gRpt_Formula_Value(3) = Format(JIGBIL, "####-##-##")


//			/// SubReport



//			/// Procedure ����"
//			sQry = "EXEC [PH_PY305_02] '" + CLTCOD + "', '" + sCode + "',  '" + StdYear + "',  '" + Quarter + "'";

//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto PH_PY305_Print_Report01_Error;
//			}

//			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "Y", "V", , 1) == false) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("gCryReport_Action : ����!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			}

//			//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			PH_PY305_Print_Report01_Error:

//			if (ErrNum == 1) {
//				//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oRecordSet = null;
//				MDC_Com.MDC_GF_Message(ref "����� �����Ͱ� �����ϴ�. Ȯ���� �ּ���.", ref "E");
//			} else {
//				//UPGRADE_NOTE: oRecordSet ��ü�� �������� �����Ǿ�� �Ҹ�˴ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oRecordSet = null;
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY305_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			}

//		}
//	}
//}

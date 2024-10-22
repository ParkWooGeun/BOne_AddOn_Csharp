﻿using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 사용자 권한 조회(화면별)
    /// </summary>
    internal class PH_PY998 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY998B;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;
        private string cSuperUserYN; //최초 화면이 실행 될때 Superuser 여부 저장

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY998.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY998_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY998");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY998_CreateItems();
                PH_PY998_ComboBox_Setting();
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
                oForm.ActiveItem = "PermID"; //최초 Load 시 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY998_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oDS_PH_PY998B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                //화면구분
                oForm.DataSources.UserDataSources.Add("PGClass", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("PGClass").Specific.DataBind.SetBound(true, "", "PGClass");

                //화면권한ID
                oForm.DataSources.UserDataSources.Add("PermID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("PermID").Specific.DataBind.SetBound(true, "", "PermID");

                //화면권한명
                oForm.DataSources.UserDataSources.Add("PermNM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("PermNM").Specific.DataBind.SetBound(true, "", "PermNM");

                //권한
                oForm.DataSources.UserDataSources.Add("Perm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Perm").Specific.DataBind.SetBound(true, "", "Perm");

                //사용자ID
                oForm.DataSources.UserDataSources.Add("UserCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("UserCode").Specific.DataBind.SetBound(true, "", "UserCode");

                //사용자성명
                oForm.DataSources.UserDataSources.Add("UserName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("UserName").Specific.DataBind.SetBound(true, "", "UserName");
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
        /// PH_PY030_ComboBox_Setting
        /// </summary>
        private void PH_PY998_ComboBox_Setting()
        {
            try
            {
                oForm.Freeze(true);

                oForm.Items.Item("PGClass").Specific.ValidValues.Add("1", "System 화면");
                oForm.Items.Item("PGClass").Specific.ValidValues.Add("2", "AddOn 화면");
                oForm.Items.Item("PGClass").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.Items.Item("Perm").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("Perm").Specific.ValidValues.Add("1", "모든 권한");
                oForm.Items.Item("Perm").Specific.ValidValues.Add("2", "읽기 전용");
                oForm.Items.Item("Perm").Specific.ValidValues.Add("3", "권한 없음");
                oForm.Items.Item("Perm").Specific.ValidValues.Add("4", "여러 권한");
                oForm.Items.Item("Perm").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
        /// System Form 권한 조회(원본)
        /// </summary>
        private void PH_PY998_SelectSystemFormPermisson()
        {
            string sQry;
            string permName = string.Empty;

            SAPbobsCOM.SBObob oSBObob = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset); //사용자ID 조회용
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset); //GetSystemPermission 저장용
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);

                /*
                    1. 현재 사용자 ID 조회, 재직중인 사원(OUSR User_Code 조회)
                    2. 1의 카운트만큼 루프
                    3. 조회하고자 하는 권한(1:모든 권한, 2:읽기 전용, 3:권한 없음)을 가진 사용자 ID를 저장(DataRow)
                    4. 저장된 DataRow의 카운트만큼 루프
                        4-1. matrix의 각 필드에 매칭 데이터 출력
                */

                string permID = oForm.Items.Item("PermID").Specific.Value.ToString().Trim();
                string perm = oForm.Items.Item("Perm").Specific.Selected.Value.ToString().Trim();
                string userCode = oForm.Items.Item("UserCode").Specific.Value.ToString().Trim();

                sQry = "EXEC PH_PY998_01 '";
                sQry += userCode + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY998B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                System.Data.DataTable userTable = this.PH_PY998_GetUserPermissionTable();
                System.Data.DataRow userRow = null;

                //조건에 맞는 데이터를 DataRow에 저장
                for (int loopCount = 0; loopCount <= oRecordSet01.RecordCount - 1; loopCount++)
                {
                    userRow = userTable.NewRow();

                    oRecordSet02 = oSBObob.GetSystemPermission(oRecordSet01.Fields.Item("UserCode").Value, permID); //권한 조회

                    if (oRecordSet02.Fields.Item(0).Value.ToString().Trim() == "1") //모든 권한(Read/Write)
                    {
                        permName = "모든 권한";
                    }
                    else if (oRecordSet02.Fields.Item(0).Value.ToString().Trim() == "2") //읽기 전용(Read Only)
                    {
                        permName = "읽기 전용";
                    }
                    else if (oRecordSet02.Fields.Item(0).Value.ToString().Trim() == "3") //권한 없음(Not authorized)
                    {
                        permName = "권한 없음";
                    }
                    else if (oRecordSet02.Fields.Item(0).Value.ToString().Trim() == "4") //여러 권한(Various authorized)
                    {
                        permName = "여러 권한";
                    }
                    else if (oRecordSet02.Fields.Item(0).Value.ToString().Trim() == "6") //정의 안됨(Not defined)
                    {
                        permName = "정의 안됨";
                    }

                    if (perm == "%") //전체
                    {
                        userRow["UserID"] = oRecordSet01.Fields.Item("UserCode").Value.ToString().Trim(); //사용자ID
                        userRow["UserName"] = oRecordSet01.Fields.Item("UserName").Value.ToString().Trim(); //성명
                        userRow["MSTCOD"] = oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim(); //사번
                        userRow["BPLName"] = oRecordSet01.Fields.Item("BPLName").Value.ToString().Trim(); //소속사업장
                        userRow["TeamName"] = oRecordSet01.Fields.Item("TeamName").Value.ToString().Trim(); //소속부서
                        userRow["Perm"] = permName; //권한
                        userTable.Rows.Add(userRow);
                    }
                    else //전체가 아니면 조회조건에 맞는 권한
                    {
                        if (oRecordSet02.Fields.Item(0).Value == perm)
                        {
                            userRow["UserID"] = oRecordSet01.Fields.Item("UserCode").Value.ToString().Trim(); //사용자ID
                            userRow["UserName"] = oRecordSet01.Fields.Item("UserName").Value.ToString().Trim(); //성명
                            userRow["MSTCOD"] = oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim(); //사번
                            userRow["BPLName"] = oRecordSet01.Fields.Item("BPLName").Value.ToString().Trim(); //소속사업장
                            userRow["TeamName"] = oRecordSet01.Fields.Item("TeamName").Value.ToString().Trim(); //소속부서
                            userRow["Perm"] = permName; //권한
                            userTable.Rows.Add(userRow);
                        }
                    }

                    oRecordSet01.MoveNext();
                }

                //Matrix에 출력
                for (int loopCount = 0; loopCount <= userTable.Rows.Count - 1; loopCount++)
                {
                    if (loopCount + 1 > oDS_PH_PY998B.Size)
                    {
                        oDS_PH_PY998B.InsertRecord(loopCount);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY998B.Offset = loopCount;

                    oDS_PH_PY998B.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));
                    oDS_PH_PY998B.SetValue("U_ColReg01", loopCount, userTable.Rows[loopCount]["UserID"].ToString()); //사용자ID
                    oDS_PH_PY998B.SetValue("U_ColReg02", loopCount, userTable.Rows[loopCount]["UserName"].ToString()); //성명
                    oDS_PH_PY998B.SetValue("U_ColReg03", loopCount, userTable.Rows[loopCount]["MSTCOD"].ToString()); //사번
                    oDS_PH_PY998B.SetValue("U_ColReg04", loopCount, userTable.Rows[loopCount]["BPLName"].ToString()); //소속사업장
                    oDS_PH_PY998B.SetValue("U_ColReg05", loopCount, userTable.Rows[loopCount]["TeamName"].ToString()); //소속부서
                    oDS_PH_PY998B.SetValue("U_ColReg06", loopCount, userTable.Rows[loopCount]["Perm"].ToString()); //권한

                    ProgBar01.Value += 1;
                    ProgBar01.Text = "조회중...!";
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, 1, "Ok", "", "");
            }
            finally
            {
                ProgBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 접속한 User가 Superuser가 아닐 경우 Superuser권한 임시 부여 (GetSystemPermission 메소드 사용권한 확보)
        /// 사용안함
        /// </summary>
        private void PH_PY998_ChangeSuperUser()
        {
            SAPbobsCOM.Users oUser = null;
            string lSuperUserYN; //메소드가 호출되는 시점의 Superuser여부를 저장할 지역변수

            int errCode = 0;
            int errNum = 0;
            string errMsg = string.Empty;

            try
            {
                oUser = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);

                if (oUser.GetByKey(PSH_Globals.oCompany.UserSignature))
                {
                    int userSign = PSH_Globals.oCompany.UserSignature;

                    lSuperUserYN = oUser.Superuser.ToString();

                    if (cSuperUserYN == "tNO" && lSuperUserYN == "tNO")
                    {
                        oUser.Superuser = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    else if (cSuperUserYN == "tNO" && lSuperUserYN == "tYES")
                    {
                        oUser.Superuser = SAPbobsCOM.BoYesNoEnum.tNO;
                    }

                    int returnValue = oUser.Update();

                    if (returnValue != 0)
                    {
                        PSH_Globals.oCompany.GetLastError(out errCode, out errMsg);
                        errNum = 1;
                        throw new Exception();
                    }
                }
            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("[" + errCode + "]" + errMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                if (oUser != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUser);
                }
            }
        }

        /// <summary>
        /// 화면이 호출되는 시점에 Superuser 여부를 클래스전역변수에 저장
        /// 사용안함
        /// </summary>
        private void PH_PY998_GetSuperUser()
        {
            SAPbobsCOM.Users oUser = null;
            
            try
            {
                oUser = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);

                if (oUser.GetByKey(PSH_Globals.oCompany.UserSignature))
                {
                    cSuperUserYN = oUser.Superuser.ToString(); //최초 Superuser 여부를 클래스전역변수에 저장
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oUser != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUser);
                }
            }
        }

        /// <summary>
        /// AddOn Form 권한 조회
        /// </summary>
        private void PH_PY998_SelectAddOnFormPermisson()
        {
            string sQry;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset); //사용자ID 조회용
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);

                string permID = oForm.Items.Item("PermID").Specific.Value.ToString().Trim();
                string userCode = oForm.Items.Item("UserCode").Specific.Value.ToString().Trim();
                string sPerm = oForm.Items.Item("Perm").Specific.Value.ToString().Trim();

                sQry = "EXEC PH_PY998_02 '";
                sQry += permID + "','";
                sQry += userCode + "','";
                sQry += sPerm + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY998B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                //Matrix에 출력
                for (int loopCount = 0; loopCount <= oRecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount + 1 > oDS_PH_PY998B.Size)
                    {
                        oDS_PH_PY998B.InsertRecord(loopCount);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY998B.Offset = loopCount;

                    oDS_PH_PY998B.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));
                    oDS_PH_PY998B.SetValue("U_ColReg01", loopCount, oRecordSet01.Fields.Item("UserCode").Value.ToString().Trim()); //사용자ID
                    oDS_PH_PY998B.SetValue("U_ColReg02", loopCount, oRecordSet01.Fields.Item("UserName").Value.ToString().Trim()); //성명
                    oDS_PH_PY998B.SetValue("U_ColReg03", loopCount, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim()); //사번
                    oDS_PH_PY998B.SetValue("U_ColReg04", loopCount, oRecordSet01.Fields.Item("BPLName").Value.ToString().Trim()); //소속사업장
                    oDS_PH_PY998B.SetValue("U_ColReg05", loopCount, oRecordSet01.Fields.Item("TeamName").Value.ToString().Trim()); //소속부서
                    oDS_PH_PY998B.SetValue("U_ColReg06", loopCount, oRecordSet01.Fields.Item("Perm").Value.ToString().Trim()); //권한

                    ProgBar01.Value += 1;
                    ProgBar01.Text = "조회중...!";

                    oRecordSet01.MoveNext();
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, 1, "Ok", "", "");
            }
            finally
            {
                ProgBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        private bool PH_PY998_DataValidCheck()
        {
            bool returnValue = false;
            int errNum = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("PermID").Specific.Value.ToString().Trim())) //화면권한ID
                {
                    errNum = 1;
                    throw new Exception();
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    dataHelpClass.MDC_GF_Message("화면권한ID는 필수사항입니다. 확인하세요.", "E");
                    oForm.Items.Item("PermID").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 권한별 사용자 정보 저장용 DataTable 생성
        /// </summary>
        /// <returns>조회조건과 일치하는 사용자 권한을 저장한 DataTable</returns>
        private System.Data.DataTable PH_PY998_GetUserPermissionTable()
        {
            System.Data.DataTable userTable = new System.Data.DataTable("UserInfo");

            //사용자ID
            System.Data.DataColumn userID = new System.Data.DataColumn();
            userID.DataType = System.Type.GetType("System.String");
            userID.ColumnName = "UserID";
            //userID.AutoIncrement = true;
            userTable.Columns.Add(userID);

            //성명
            System.Data.DataColumn userName = new System.Data.DataColumn();
            userName.DataType = System.Type.GetType("System.String");
            userName.ColumnName = "UserName";
            //userName.DefaultValue = "Fname";
            userTable.Columns.Add(userName);

            //사번
            System.Data.DataColumn mstCode = new System.Data.DataColumn();
            mstCode.DataType = System.Type.GetType("System.String");
            mstCode.ColumnName = "MSTCOD";
            userTable.Columns.Add(mstCode);

            //소속사업장
            System.Data.DataColumn bplName = new System.Data.DataColumn();
            bplName.DataType = System.Type.GetType("System.String");
            bplName.ColumnName = "BPLName";
            userTable.Columns.Add(bplName);

            //소속부서
            System.Data.DataColumn teamName = new System.Data.DataColumn();
            teamName.DataType = System.Type.GetType("System.String");
            teamName.ColumnName = "TeamName";
            userTable.Columns.Add(teamName);

            //권한
            System.Data.DataColumn perm = new System.Data.DataColumn();
            perm.DataType = System.Type.GetType("System.String");
            perm.ColumnName = "Perm";
            userTable.Columns.Add(perm);

            //Create an array for DataColumn objects.
            System.Data.DataColumn[] keys = new System.Data.DataColumn[1];
            keys[0] = userID;
            userTable.PrimaryKey = keys;

            return userTable;
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
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
        
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "BtnSearch")
                    {
                        if (PH_PY998_DataValidCheck() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }

                        if (oForm.Items.Item("PGClass").Specific.Selected.Value == "1")
                        {
                            this.PH_PY998_SelectSystemFormPermisson();
                        }
                        else
                        {
                            this.PH_PY998_SelectAddOnFormPermisson();
                        }
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY998B);
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
                if (pVal.BeforeAction == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "PermID", ""); //화면권한ID
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "UserCode", ""); //사용자ID
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string query;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                        }
                        else
                        {
                            if (pVal.ItemUID == "PermID")
                            {
                                if (oForm.Items.Item("PGClass").Specific.Selected.Value == "1") //System Form
                                {
                                    oForm.Items.Item("PermNM").Specific.Value = dataHelpClass.Get_ReData("U_KorName", "U_PermID", "[@PSH_PERMISSION_ID]", "'" + oForm.Items.Item("PermID").Specific.Value + "'", ""); //화면권한명
                                }
                                else //AddOn Form
                                {
                                    oForm.Items.Item("PermNM").Specific.Value = dataHelpClass.Get_ReData("String", "UniqueID", "[Authority_Screen]", "'" + oForm.Items.Item("PermID").Specific.Value + "'", ""); //화면권한명
                                }
                            }
                            else if (pVal.ItemUID == "UserCode")
                            {
                                query = "  SELECT    T1.U_FullName";
                                query += " FROM      OUSR AS T0";
                                query += "           LEFT JOIN";
                                query += "           OHEM AS T1";
                                query += "               ON T0.USERID = T1.UserID";
                                query += " WHERE T0.USER_CODE = '" + oForm.Items.Item("UserCode").Specific.Value + "'";

                                oRecordSet01.DoQuery(query);

                                oForm.Items.Item("UserName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim(); //사용자성명
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oMat01.SelectRow(pVal.Row, true, false);

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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
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

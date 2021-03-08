using System;
using System.Collections.Generic;
using System.Reflection;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 이벤트 필터
    /// 1. 현재 Namespace의 전체 클래스 조회
    /// 2. 클래스의 메소드 중 해당 이벤트 필터 메소드가 존재하면 필터 추가
    /// </summary>
    internal static class PSH_SetFilter
    {
        private static List<Type> classList = GetClasses("PSH_BOne_AddOn_Class"); //Namespace 내의 모든 클래스 조회

        public static void Execute()
		{
			SAPbouiCOM.EventFilters oFilters = null;
			SAPbouiCOM.EventFilter oFilter = null;

			oFilters = new SAPbouiCOM.EventFilters();

            ITEM_PRESSED(ref oFilter, ref oFilters); //1
            KEY_DOWN(ref oFilter, ref oFilters); //2
            GOT_FOCUS(ref oFilter, ref oFilters); //3
            LOST_FOCUS(ref oFilter, ref oFilters); //4
            COMBO_SELECT(ref oFilter, ref oFilters); //5
            CLICK(ref oFilter, ref oFilters); //6
            DOUBLE_CLICK(ref oFilter, ref oFilters); //7
            MATRIX_LINK_PRESSED(ref oFilter, ref oFilters); //8
            MATRIX_COLLAPSE_PRESSED(ref oFilter, ref oFilters); //9
            VALIDATE(ref oFilter, ref oFilters); //10
            MATRIX_LOAD(ref oFilter, ref oFilters); //11
            DATASOURCE_LOAD(ref oFilter, ref oFilters); //12
            FORM_LOAD(ref oFilter, ref oFilters); //16
            FORM_UNLOAD(ref oFilter, ref oFilters); //17
            FORM_ACTIVATE(ref oFilter, ref oFilters); //18
            FORM_DEACTIVATE(ref oFilter, ref oFilters); //19
            FORM_CLOSE(ref oFilter, ref oFilters); //20
            FORM_RESIZE(ref oFilter, ref oFilters); //21
            FORM_KEY_DOWN(ref oFilter, ref oFilters); //22
            FORM_MENU_HILIGHT(ref oFilter, ref oFilters); //23
            PRINT(ref oFilter, ref oFilters); //24
            PRINT_DATA(ref oFilter, ref oFilters); //25
            CHOOSE_FROM_LIST(ref oFilter, ref oFilters); //27
            RIGHT_CLICK(ref oFilter, ref oFilters); //28
            MENU_CLICK(ref oFilter, ref oFilters); //32
            FORM_DATA_ADD(ref oFilter, ref oFilters); //33
            FORM_DATA_UPDATE(ref oFilter, ref oFilters); //34
            FORM_DATA_DELETE(ref oFilter, ref oFilters); //35
            FORM_DATA_LOAD(ref oFilter, ref oFilters); //36
            
            PSH_Globals.SBO_Application.SetFilter(oFilters);
			
			oFilter = null;
			oFilters = null;
		}

        private static void ITEM_PRESSED(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_ITEM_PRESSED", classList[i], ref oFilter);
            }
        }

        private static void KEY_DOWN(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_KEY_DOWN", classList[i], ref oFilter);
            }
        }

        private static void GOT_FOCUS(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_GOT_FOCUS", classList[i], ref oFilter);
            }
        }

        private static void LOST_FOCUS(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_LOST_FOCUS", classList[i], ref oFilter);
            }
        }

        private static void COMBO_SELECT(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_COMBO_SELECT", classList[i], ref oFilter);
            }
        }

        private static void CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_CLICK", classList[i], ref oFilter);
            }
        }

        private static void DOUBLE_CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_DOUBLE_CLICK", classList[i], ref oFilter);
            }
        }

        private static void MATRIX_LINK_PRESSED(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_MATRIX_LINK_PRESSED", classList[i], ref oFilter);
            }
        }

        private static void MATRIX_COLLAPSE_PRESSED(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_MATRIX_COLLAPSE_PRESSED", classList[i], ref oFilter);
            }
        }

        private static void VALIDATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_VALIDATE", classList[i], ref oFilter);
            }
        }

        private static void MATRIX_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_MATRIX_LOAD", classList[i], ref oFilter);
            }
        }

        private static void DATASOURCE_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_DATASOURCE_LOAD", classList[i], ref oFilter);
            }
        }

        private static void FORM_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_LOAD", classList[i], ref oFilter);
            }
        }

        private static void FORM_UNLOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_UNLOAD", classList[i], ref oFilter);
            }
        }

        private static void FORM_ACTIVATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_ACTIVATE", classList[i], ref oFilter);
            }
        }

        private static void FORM_DEACTIVATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_DEACTIVATE", classList[i], ref oFilter);
            }
        }

        private static void FORM_CLOSE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_CLOSE", classList[i], ref oFilter);
            }
        }

        private static void FORM_RESIZE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_RESIZE", classList[i], ref oFilter);
            }
        }

        private static void FORM_KEY_DOWN(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_KEY_DOWN", classList[i], ref oFilter);
            }
        }

        private static void FORM_MENU_HILIGHT(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_MENU_HILIGHT", classList[i], ref oFilter);
            }
        }

        private static void PRINT(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_PRINT);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_PRINT", classList[i], ref oFilter);
            }
        }

        private static void PRINT_DATA(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_PRINT_DATA);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_PRINT_DATA", classList[i], ref oFilter);
            }
        }

        private static void CHOOSE_FROM_LIST(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_CHOOSE_FROM_LIST", classList[i], ref oFilter);
            }
        }

        private static void RIGHT_CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_RIGHT_CLICK", classList[i], ref oFilter);
            }
        }

        private static void MENU_CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_MENU_CLICK", classList[i], ref oFilter);
            }
        }

        private static void FORM_DATA_ADD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_DATA_ADD", classList[i], ref oFilter);
            }
        }

        private static void FORM_DATA_UPDATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_DATA_UPDATE", classList[i], ref oFilter);
            }
        }

        private static void FORM_DATA_DELETE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_DATA_DELETE", classList[i], ref oFilter);
            }
        }

        private static void FORM_DATA_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
        {
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);

            for (int i = 0; i < classList.Count; i++)
            {
                RegisterEventFilter("Raise_EVENT_FORM_DATA_LOAD", classList[i], ref oFilter);
            }
        }

        /// <summary>
        /// Namespcae의 전체 클래스 조회
        /// </summary>
        /// <param name="nameSpace">Namespace</param>
        /// <returns></returns>
        private static List<Type> GetClasses(string nameSpace)
        {
            List<Type> typeList = new List<Type>();

            foreach (Type t in Assembly.GetExecutingAssembly().GetTypes())
            {
                if (t.Namespace == nameSpace && !t.IsAbstract) //추상클래스, 인터페이스 제외
                {
                    typeList.Add(t);
                }
            }

            return typeList;
        }

        /// <summary>
        /// SAP B1 이벤트 필터 등록
        /// </summary>
        /// <param name="eventMethodName">이벤트 필터 메소드 명</param>
        /// <param name="classType">클래스 타입</param>
        /// <param name="eventFilter">이벤트 필터</param>
        private static void RegisterEventFilter(string eventMethodName, Type classType, ref SAPbouiCOM.EventFilter eventFilter)
        {
            try
            {
                MethodInfo[] arrayMethodInfo = classType.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly);

                for (int i = 0; i < arrayMethodInfo.Length; i++)
                {
                    if (arrayMethodInfo[i].Name == eventMethodName)
                    {
                        eventFilter.AddEx(classType.Name);
                    }
                }
            }
            catch (System.Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
    }
}

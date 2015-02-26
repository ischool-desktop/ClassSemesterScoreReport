using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassSemesterScoreReport
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["班級", "資料統計"];
            item1["報表"]["成績相關報表"]["班級學期成績單"].Enable = false;
            item1["報表"]["成績相關報表"]["班級學期成績單"].Click += delegate
            {
                new MainForm().ShowDialog();
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                item1["報表"]["成績相關報表"]["班級學期成績單"].Enable = K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 && Permissions.班級學期成績單權限;
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["班級"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.班級學期成績單, "班級學期成績單"));

        }
    }
}

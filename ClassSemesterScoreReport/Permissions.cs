using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClassSemesterScoreReport
{
    class Permissions
    {
        public static bool 班級學期成績單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級學期成績單].Executable;
            }
        }

        public static string 班級學期成績單 = "ClassSemesterScoreReport-{0183C5AB-BD58-4468-BBC6-D0AD48993859}";
    }
}

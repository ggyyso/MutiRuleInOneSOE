using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MultiRuleInOneSOE
{
    /// <summary>
    /// 用地性质
    /// </summary>
  public class ydxzClass
  {
      /// <summary>
      /// 用地性质
      /// </summary>
      private string type;

      public string Type
      {
          get { return type; }
          set { type = value; }
      }
      /// <summary>
      /// 用地代码
      /// </summary>
      private string code;

      public string Code
      {
          get { return code; }
          set { code = value; }
      }
      /// <summary>
      /// 控制线检测项目
      /// </summary>
      private List<jcxClass> kzxjc;

      public List<jcxClass> Kzxjc
      {
          get { return kzxjc; }
          set { kzxjc = value; }
      }
      /// <summary>
      /// 土地类型检测项
      /// </summary>
      private List<jcxClass> tdlyjc;

      public List<jcxClass> Tdlyjc
      {
          get { return tdlyjc; }
          set { tdlyjc = value; }
      }

      /// <summary>
      /// 矿产检测项
      /// </summary>
      private List<jcxClass> kchgjc;

      public List<jcxClass> Kchgjc
      {
          get { return kchgjc; }
          set { kchgjc = value; }
      }



  }
    /// <summary>
    /// 检测项
    /// </summary>
  public class jcxClass
  {
      /// <summary>
      /// 检测类型
      /// </summary>
      private string type;

      public string Type
      {
          get { return type; }
          set { type = value; }
      }
      /// <summary>
      /// 检测条件
      /// </summary>
      private bool condi;

      public bool Condi
      {
          get { return condi; }
          set { condi = value; }
      }


  }

  public class resultClass
  {
      private double area;

      public double Area
      {
          get { return area; }
          set { area = value; }
      }
      private string result;

      public string Result
      {
          get { return result; }
          set { result = value; }
      }

  }
}

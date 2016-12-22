using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    public partial class WorkFlow
    {

        public virtual void fromimage<T1>(Action<T1> ac, T1 t1) {
            System.Console.Write("我要打印字在屏幕上");
            ac(t1);
            System.Console.Write("打印字在屏幕上完成");
        }
        public virtual void fromimage<T1,T2>(Action<T1,T2> ac, T1 t1,T2 t2) {
            System.Console.Write("我要打印字在屏幕上");
            ac(t1, t2);
            System.Console.Write("打印字在屏幕上完成");
        }
        public virtual void fromimage<T1,T2,T3>(Action<T1,T2,T3> ac, T1 t1,T2 t2,T3 t3)  {
            System.Console.Write("我要打印字在屏幕上");
            ac(t1, t2,t3);
            System.Console.Write("打印字在屏幕上完成");
        }
        public virtual void fromimage<T1,T2,T3,T4>(Action<T1,T2,T3,T4> ac, T1 t1, T2 t2, T3 t3, T4 t4 ) {
            System.Console.Write("我要打印字在屏幕上");
            ac(t1, t2,t3,t4);
            System.Console.Write("打印字在屏幕上完成");
        }
        
        public virtual void StartVersion<T>(Action<T> ac, T t1) 
        {
            System.Console.Write("我要生成一个图片");
            this.fromimage<T>(ac, t1);
            System.Console.Write("生成图生完成");
        }

        public virtual void StartVersion< T1,T2>(Action<T1,T2> ac, T1 t1, T2 t2)
        {
            System.Console.Write("我要生成一个图片");
            fromimage<T1,T2>(ac, t1, t2);
            System.Console.Write("生成图生完成");
            
        }

        public virtual void StartVersion< T1,T2,T3>(Action<T1,T2,T3>ac, T1 t1, T2 t2,T3 t3)
        {
            System.Console.Write("我要生成一个图片");
            fromimage<T1, T2,T3>(ac, t1, t2,t3);
            System.Console.Write("生成图生完成");
        }

        public virtual void StartVersion<T1, T2,T3,T4>(Action<T1,T2,T3,T4> ac, T1 t1, T2 t2, T3 t3, T4 t4)
        {
            System.Console.Write("我要生成一个图片");
            fromimage< T1, T2, T3,T4>(ac, t1, t2, t3,t4);
            System.Console.Write("生成图生完成");
        }



    }
}

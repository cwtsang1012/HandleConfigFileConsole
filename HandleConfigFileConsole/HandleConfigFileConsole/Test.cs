using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    public class Test
    {
        private string _className;
        public string SessionId { get; set; }
        public string Description { get; set; }
        public string Path { get; set; }

        public Test() 
        {
            _className = this.GetType().Name;
        }

        public Test(string sessionId, string description, string path) 
        {
            _className = this.GetType().Name;
            SessionId = sessionId;
            Description = description;
            Path = path;
        }

        public virtual void Print() 
        {
            Console.WriteLine("---------" +_className + "---------");
            Console.WriteLine("Session Id: " + SessionId);
            Console.WriteLine("Description: " + Description);
        }
    }

    public class Test001 : Test 
    {
        public Test001() { }
        public Test001(string sessionId, string description, string path) : base(sessionId, description, path) { }

        public override void Print()
        {
            base.Print();
            createTokFile();
        }
        
        public void createTokFile() 
        {
            Console.WriteLine("Token file is created");
        }
    }

    public class Test002 : Test
    {
        public Test002() { }
        public Test002(string sessionId, string description, string path) : base(sessionId, description, path) { }
        
        public override void Print()
        {
            base.Print();
            copyTokFileFr01();
        }

        public void copyTokFileFr01()
        {
            Console.WriteLine("Token file is copied from Test001");
        }
    }

    public class Test003 : Test
    {
        public Test003() { }
        public Test003(string sessionId, string description, string path) : base(sessionId, description, path) { }
        
        public override void Print()
        {
            base.Print();
            createTxtFile();
        }
        public void createTxtFile()
        {
            Console.WriteLine("Txt file is created");
        }
    }

    public class Test004 : Test
    {
       public Test004() { }
       public Test004(string sessionId, string description, string path) : base(sessionId, description, path) { }
    }

}

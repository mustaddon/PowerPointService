using SimpleInjector;
using System;
using System.Collections.Generic;
using System.Text;
using RandomSolutions;
using System.Linq;

namespace NetConsoleApp
{
    class IoC
    {
        public static Container Container = _create();
        
        static Container _create()
        {
            var container = new Container();

            container.Register<PowerPointService>();
            container.Collection.Register<IPipeTransform>(typeof(IPipeTransform).Assembly);

            container.Verify();
            return container;
        }
    }
}

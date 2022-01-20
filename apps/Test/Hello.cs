using Kompas6API5;

namespace Test
{
    // ReSharper disable once UnusedType.Global
    public class Hello : BaseCommand
    {
        public Hello() : base("Hello - the first kompas app")
        {
        }

        protected override void Action(short command, short mode, object kompasObj)
        {
            KompasObject kompas = (KompasObject) kompasObj;
            kompas.ksMessage("Hello zzz Kompas!");
        }
    }
}
﻿using Kompas6API5;
using Base;

namespace RunCommands
{
    /// <summary>
    /// Create new part with english namings
    /// </summary>
    ///
    // ReSharper disable once UnusedType.Global
    public class CreatePart : BaseCommand
    {
        public CreatePart() : base("API7 3D test")
        {
        }

        protected override void Action(short command, short mode, object kompasObj)
        {
            KompasObject kompas = (KompasObject)kompasObj;
            bool isSuccess;
            do
            {
                // ReSharper disable once NotAccessedVariable
                Document3D d3d;
                isSuccess = DocHelpers.CreateDoc3D(out d3d, kompas, DocHelpers.Doc3DType.Part);

            } while (false);

            if (isSuccess)
            {
                kompas.ksMessage("");
            }
            else
            {
                kompas.ksMessage("Hello Kompas!");
            }
        }
    }
}
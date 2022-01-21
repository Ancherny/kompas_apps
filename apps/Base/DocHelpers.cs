using JetBrains.Annotations;
using Kompas6API5;
using KAPITypes;
using Kompas6Constants;
using Kompas6Constants3D;

namespace Base;

public static class DocHelpers
{
    public enum Doc3DType
    {
        Part,
        Assembly
    }

    public static void CreateDoc2D(object kompasObj)
    {
        var kompas = (KompasObject)kompasObj;

        var dp = (DocumentParam)kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);;
        dp.Init();
        dp.type = (short)DocType.lt_DocPart3D;

        var d2d = (Document2D)kompas.Document2D();
        d2d.ksCreateDocument(dp);
        kompas.ksMessage("Doc2D created!");
    }

    public static bool CreateDoc3D(out Document3D d3d, [NotNull] KompasObject kompas, Doc3DType doc3DType)
    {
        d3d = null;
        bool isSuccess;
        d3d = (Document3D)kompas.Document3D();
        switch (doc3DType)
        {
            case Doc3DType.Part:
                isSuccess = d3d.Create(false, true);
                break;
            case Doc3DType.Assembly:
                isSuccess = d3d.Create(false, false);
                break;
            default:
                isSuccess = false;
                break;
        }
        return isSuccess;
    }

    public static void SetPartFirstEntityName([NotNull] ksPart part, [NotNull] string name, Obj3dType entityType)
    {
        var collection = (ksEntityCollection)part.EntityCollection((short)entityType);
        var entity = (ksEntity)collection.GetByIndex(0);
        entity.name = name;
        entity.Update();
    }
}
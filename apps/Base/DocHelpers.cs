using System;
using JetBrains.Annotations;
using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;
using KompasAPI7;

namespace Base
{
    public static class DocHelpers
    {
        private delegate IModelObject GetModelObject(ksObj3dTypeEnum obj3d);

        public static bool CreateNew(KompasObject kompas, DocumentTypeEnum docType)
        {
            bool isSuccess = true;
            do
            {
                string newName;
                var app = (IApplication)kompas.ksGetApplication7();
                if (app == null)
                {
                    isSuccess = false;
                    break;
                }
                var d3d = (IKompasDocument3D)app.Documents.Add(docType);
                if (d3d == null)
                {
                    isSuccess = false;
                    break;
                }

                // ReSharper disable once SwitchStatementHandlesSomeKnownEnumValuesWithDefault
                switch (docType)
                {
                    case DocumentTypeEnum.ksDocumentPart:
                        newName = "Part";
                        break;
                    case DocumentTypeEnum.ksDocumentAssembly:
                        newName = "Assembly";
                        break;
                    default:
                        newName = docType.ToString();
                        break;
                }

                var topPart = d3d.TopPart;
                topPart.Name = newName;
                topPart.Update();
                RenamePartOrigin(topPart);

            } while (false);

            return isSuccess;
        }

        // ReSharper disable once MemberCanBePrivate.Global
        public static T TryGetObjectOfType<T>(out bool isSuccess, [NotNull] object genericObj) where T : IKompasAPIObject
        {
            isSuccess = true;
            T kompasApiObj = default;
            try
            {
                kompasApiObj = (T)genericObj;
            }
            catch (Exception)
            {
                isSuccess = false;
            }
            return kompasApiObj;
        }

        public static bool RenameSelectedPart(KompasObject kompas, [NotNull] string name)
        {
            bool isSuccess;
            do
            {
                var app = (IApplication)kompas.ksGetApplication7();
                if (app == null)
                {
                    isSuccess = false;
                    break;
                }
                var doc = (IKompasDocument3D)app.ActiveDocument;
                if (doc == null)
                {
                    isSuccess = false;
                    break;
                }

                SelectionManager sm = doc.SelectionManager;
                var so = sm.SelectedObjects;
                if (so == null)
                {
                    kompas.ksMessage("No editable objects selected.");
                    isSuccess = false;
                    break;
                }

                if (so.GetType().IsArray)
                {
                    kompas.ksMessage("Multiple selected objects are not supported.");
                    isSuccess = false;
                    break;
                }

                var part = TryGetObjectOfType<IPart7>(out isSuccess, so);
                if (isSuccess)
                {
                    part.Name = name;
                    part.Update();
                    RenamePartOrigin(part);
                }
                else
                {
                    var origin = TryGetObjectOfType<ILocalCoordinateSystem>(out isSuccess, so);
                    if (isSuccess)
                    {
                        origin.Name = name;
                        origin.Update();
                        RenameOriginComponents(obj3d => origin.DefaultObject[obj3d]);
                    }
                }

            } while (false);

            return isSuccess;
        }

        private static void RenamePartOrigin([NotNull] IPart7 part)
        {
            IModelObject GetModelObject(ksObj3dTypeEnum obj3d) => part.DefaultObject[obj3d];
            RenameOriginComponent("Origin", ksObj3dTypeEnum.o3d_pointCS, GetModelObject);
            RenameOriginComponents(GetModelObject);
        }

        private static void RenameOriginComponent(
            [NotNull] string name,
            ksObj3dTypeEnum originComponentType,
            GetModelObject getModelObject)
        {
            var modelObject = getModelObject(originComponentType);
            modelObject.Name = name;
            modelObject.Update();
        }

        private static void RenameOriginComponents(GetModelObject getModelObject)
        {
            RenameOriginComponent("Plane_XY", ksObj3dTypeEnum.o3d_planeXOY, getModelObject);
            RenameOriginComponent("Plane_XZ", ksObj3dTypeEnum.o3d_planeXOZ, getModelObject);
            RenameOriginComponent("Plane_YZ", ksObj3dTypeEnum.o3d_planeYOZ, getModelObject);
            RenameOriginComponent("Axis_X", ksObj3dTypeEnum.o3d_axisOX, getModelObject);
            RenameOriginComponent("Axis_Y", ksObj3dTypeEnum.o3d_axisOY, getModelObject);
            RenameOriginComponent("Axis_Z", ksObj3dTypeEnum.o3d_axisOZ, getModelObject);
        }
    }
}

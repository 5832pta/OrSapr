

using Kompas6API5;
using System;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6Constants3D;
using reference = System.Int32;


namespace Steps.NET
{
    /// <summary>
    /// �������� ����� �������. 
    /// ����� ��������� ������ �������� ������ � ����� ����� �� �������.
    /// </summary>
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class GearPlugin
    {
        /// <summary>
        /// ��������� API ������
        /// </summary>
        private KompasObject _kompas;

        /// <summary>
        /// ��������� Automation ���������-������
        /// </summary>
        private ksDocument3D _doc3;

        /// <summary>
        /// ��������� ������������ ��������� ������� ������
        /// </summary>
        private ksDocument2D _doc;


        // ��� ����������
        /// <summary>
        /// ��������� ����� ����������.
        /// </summary>
        /// <returns>��� ����������</returns>
        [return: MarshalAs(UnmanagedType.BStr)]
        public string GetLibraryName()
        {
            return "������";
        }

        /// <summary>
        /// �������� ������� ����������
        /// </summary>
        /// <param name="command">����� ������� � ����</param>
        /// <param name="mode">����� ������</param>
        /// <param name="kompas">��������� KompasObject</param>
        public void ExternalRunCommand([In] short command, [In] short mode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object kompas)
        {
            _kompas = (KompasObject) kompas;
            if (_kompas != null)
            {
                _doc3 = (ksDocument3D) _kompas.ActiveDocument3D();
                if (_doc3 == null || _doc3.reference == 0)
                {
                    _doc3 = (ksDocument3D) _kompas.Document3D();
                    _doc3.Create(true, true);

                    _doc3.comment = "�������� ������";
                    _doc3.drawMode = 3;
                    _doc3.perspective = true;
                    _doc3.UpdateDocumentParam();
                }
                switch (command)
                {
                    case 1:
                        PluginFrm.Instanse.Plugin = this;
                        PluginFrm.Instanse.ShowDialog();
                        break;
                }
            }
        }


        /// <summary>
        /// ��������� ������ ���� ��� �������� ���� � ���� �����
        /// </summary>
        /// <param name="number">������� �����</param>
        /// <param name="itemType">��� ������</param>
        /// <param name="command">����� �������</param>
        /// <returns>������ ����</returns>
        [return: MarshalAs(UnmanagedType.BStr)]
        public string ExternalMenuItem(short number, ref short itemType, ref short command)
        {
            string result = string.Empty; //�� ��������� - ������ ������
            itemType = 1; //MENUITEM

            switch (number)
            {
                case 1:
                    result = "������� ��������";
                    command = 1;
                    break;
                case 2:
                    command = -1;
                    itemType = 3; // "ENDMENU"
                    break;
            }

            return result;
        }

        /// <summary>
        /// ������� �������� � �������
        /// </summary>
        /// <param name="angle">���� � ��������</param>
        /// <returns>���� � ��������</returns>
        private static double DegToRad(double angle)
        {
            return Math.PI * angle / 180.0;
        }

        /// <summary>
        /// �������� ������ ��������
        /// </summary>
        /// <param name="properties">������������ � ����� ���������</param>
        public void CreateModel(Gear properties)
        {
            int teethCount = properties.TeethCount;
            var angle = properties.Angle;
            var thickness = properties.Thickness;
            var shaftDiam = properties.ShaftDiam;
            var diameterOut = properties.DiameterOut;
            var diameterIn = properties.DiameterIn;
            var diameterPitch = (diameterOut + diameterIn) / 2; // ����������� ������� ������
            var alfa1 = 0.0;
            var alfa2 = 0.0;
            // ���������� ������������� ����������
            ksPart iPart = _doc3.GetPart((int) Part_Type.pNew_Part) as ksPart; //����� ��������
            if (iPart != null)
            {
                // ���������� ������������� ����������
                ksEntity planeXoy = iPart.GetDefaultEntity((short) Obj3dType.o3d_planeXOY) as ksEntity;
                ksEntity planeXoz = iPart.GetDefaultEntity((short) Obj3dType.o3d_planeXOZ) as ksEntity;
                ksEntity planeYoz = iPart.GetDefaultEntity((short) Obj3dType.o3d_planeYOZ) as ksEntity;
                // ��������� ������ (�������� ������� ������� ������)
                ksEntity iSketchEntity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketchEntity != null)
                {
                    // ��������� ���������� ������
                    ksSketchDefinition iSketchDef = iSketchEntity.GetDefinition() as ksSketchDefinition;
                    if (iSketchDef != null)
                    {
                        if (planeXoy != null)
                        {
                            // ������������� ���������,
                            // �� ������� ��������� �����
                            iSketchDef.SetPlane(planeXoy);

                            iSketchEntity.Create();

                            // ��������� ������� �������������� ������
                            // doc � ��������� �� ��������� ksDocument2D
                            _doc = iSketchDef.BeginEdit() as ksDocument2D;
                            if (_doc != null)
                            {
                                _doc.ksLineSeg(-thickness / 2, 00, thickness / 2, 0, 3); //3 - ������ �����
                                _doc.ksLineSeg(thickness / 2, shaftDiam / 2, -thickness / 2, shaftDiam / 2, 1);
                                _doc.ksLineSeg(thickness / 2, shaftDiam / 2, thickness / 2, diameterOut / 2, 1);
                                _doc.ksLineSeg(-thickness / 2, diameterOut / 2, -thickness / 2, shaftDiam / 2, 1);
                                _doc.ksLineSeg(-thickness / 2, diameterOut / 2, thickness / 2, diameterOut / 2, 1);
                            }
                            iSketchDef.EndEdit();
                        }
                    }
                }
                // ��������� ������� �������� ��������
                ksEntity iBaseRotatedEntity = iPart.NewEntity((short) Obj3dType.o3d_baseRotated) as ksEntity;
                if (iBaseRotatedEntity != null)
                {
                    // ��������� ���������� ��������
                    ksBaseRotatedDefinition iBaseRotatedDef =
                        iBaseRotatedEntity.GetDefinition() as ksBaseRotatedDefinition;
                    if (iBaseRotatedDef != null)
                    {
                        // ��������� ���������� ��������
                        iBaseRotatedDef.SetThinParam(false, (short) Direction_Type.dtNormal, 1, 1);
                        iBaseRotatedDef.SetSideParam(true, 360.0);
                        iBaseRotatedDef.toroidShapeType = false;
                        iBaseRotatedDef.SetSketch(iSketchEntity);
                        // ������� �������� ��������
                        // ��������� � ��������� ��������� ������
                        iBaseRotatedEntity.Create();
                    }
                }
                ksEntity iSketch1Entity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketch1Entity != null)
                {
                    ksSketchDefinition iSketch1Def = iSketch1Entity.GetDefinition() as ksSketchDefinition;
                    if (iSketch1Def != null)
                    {
                        if (planeYoz != null)
                        {
                            // ��������� ����� �� ��������� XOY
                            iSketch1Def.SetPlane(planeXoy);
                            iSketch1Entity.Create();
                            _doc = iSketch1Def.BeginEdit() as ksDocument2D;
                            if (_doc != null)
                            {
                                var width = properties.KeywayWidth;
                                _doc.ksLineSeg(thickness / 2, -width / 2, -thickness / 2, -width / 2, 1);
                                _doc.ksLineSeg(-thickness / 2, width / 2, -thickness / 2, -width / 2, 1);
                                _doc.ksLineSeg(-thickness / 2, width / 2, thickness / 2, width / 2, 1);
                                _doc.ksLineSeg(thickness / 2, width / 2, thickness / 2, -width / 2, 1);
                            }
                            iSketch1Def.EndEdit();
                        }
                    }
                }
                // ��������� �������� �������� �������������
                ksEntity iCutExtrusion = iPart.NewEntity((short) Obj3dType.o3d_cutExtrusion) as ksEntity;
                if (iCutExtrusion != null)
                {
                    // ��������� ���������� ���������
                    ksCutExtrusionDefinition iCutExtrusionDef =
                        iCutExtrusion.GetDefinition() as ksCutExtrusionDefinition;
                    if (iCutExtrusionDef != null)
                    {
                        // ��������� ���������� 
                        iCutExtrusionDef.SetSketch(iSketch1Entity);
                        // �����������
                        iCutExtrusionDef.directionType = (short) Direction_Type.dtNormal;
                        // �������� ��������� �� ������� �� �����������
                        iCutExtrusionDef.SetSideParam(true, (short) ksEndTypeEnum.etBlind,
                            properties.KeywayDepth + shaftDiam / 2, 0, false);
                        iCutExtrusionDef.SetThinParam(false, 0, 0, 0);
                        iCutExtrusion.Create();
                    }
                }
                // ��������� ��������� ���������
                ksEntity iOffsetPlaneEntity = iPart.NewEntity((short) Obj3dType.o3d_planeOffset) as ksEntity;
                if (iOffsetPlaneEntity != null)
                {
                    // ��������� ���������� ��������� ���������
                    ksPlaneOffsetDefinition iOffsetPlaneDef =
                        iOffsetPlaneEntity.GetDefinition() as ksPlaneOffsetDefinition;
                    if (iOffsetPlaneDef != null)
                    {
                        // ��������, ������� ��������� � ������ ��������� ��������
                        iOffsetPlaneDef.offset = thickness / 2;
                        iOffsetPlaneDef.SetPlane(planeYoz);
                        iOffsetPlaneDef.direction = false;
                        // ������ ��������� �������
                        iOffsetPlaneEntity.hidden = true;
                        // ������� ��������������� ���������
                        iOffsetPlaneEntity.Create();
                    }
                }
                // ����� ������� ������ ����� �������
                ksEntity iSketch2Entity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketch2Entity != null)
                {
                    ksSketchDefinition iSketch2Def = iSketch2Entity.GetDefinition() as ksSketchDefinition;
                    if (iSketch2Def != null)
                    {
                        // ������� ��������� � ��������������� iOffsetPlaneEntity
                        iSketch2Def.SetPlane(iOffsetPlaneEntity);
                        iSketch2Entity.Create();
                        _doc = iSketch2Def.BeginEdit() as ksDocument2D;
                        alfa1 = 360.0 / teethCount;
                        _doc.ksMtr(0, 0, 90, 1, 1);
                        // ������������ ����������� ������
                        // ������ ��������� ��� ��������
                        // ����� ������� ���� �� ���� ������
                        _doc.ksArcBy3Points(-(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(0)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(0)),
                            -diameterPitch / 2 * Math.Sin(DegToRad(alfa1 / 8)),
                            -diameterPitch / 2 * Math.Cos(DegToRad(alfa1 / 8)),
                            -diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 4)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 4)), 1);
                        _doc.ksArcByPoint(0, 0, diameterIn / 2, -diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 4)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 4)),
                            -diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 2)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 2)), -1, 1);
                        _doc.ksArcBy3Points(-diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 2)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 2)),
                            -diameterPitch / 2 * Math.Sin(DegToRad(0.625 * alfa1)),
                            -diameterPitch / 2 * Math.Cos(DegToRad(0.625 * alfa1)),
                            -(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1)), 1);
                        _doc.ksArcBy3Points(-(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1)),
                            -(diameterOut / 2 + 2) * Math.Sin(DegToRad(0.375 * alfa1)),
                            -(diameterOut / 2 + 2) * Math.Cos(DegToRad(0.375 * alfa1)),
                            -(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(0)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(0)), 1);
                        _doc.ksDeleteMtr();
                        iSketch2Def.EndEdit();
                    }
                }
                // ��������� ������� ������ ������ ����� �������
                ksEntity iSketch3Entity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketch3Entity != null)
                {
                    ksSketchDefinition iSketch3Def = iSketch3Entity.GetDefinition() as ksSketchDefinition;
                    if (iSketch3Def != null)
                    {
                        // ������ �� ��������� YOZ
                        iSketch3Def.SetPlane(planeYoz);
                        iSketch3Entity.Create();
                        _doc = iSketch3Def.BeginEdit() as ksDocument2D;
                        alfa2 = -(180 / Math.PI * (thickness * Math.Tan(Math.PI * (angle) / 180) / diameterPitch));
                        _doc.ksMtr(0, 0, 90, 1, 1);
                        // ������������ ����������� ������
                        // ������ ��������� ��� ��������
                        // ����� ������� ���� �� ���� ������
                        _doc.ksArcBy3Points(-(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(alfa2)),
                            -diameterPitch / 2 * Math.Sin(DegToRad(alfa1 / 8 + alfa2)),
                            -diameterPitch / 2 * Math.Cos(DegToRad(alfa1 / 8 + alfa2)),
                            -diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 4 + alfa2)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 4 + alfa2)), 1);
                        _doc.ksArcByPoint(0, 0, diameterIn / 2, -diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 4 + alfa2)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 4 + alfa2)),
                            -diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 2 + alfa2)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 2 + alfa2)), -1, 1);
                        _doc.ksArcBy3Points(-diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 2 + alfa2)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 2 + alfa2)),
                            -diameterPitch / 2 * Math.Sin(DegToRad(0.625 * alfa1 + alfa2)),
                            -diameterPitch / 2 * Math.Cos(DegToRad(0.625 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1 + alfa2)), 1);
                        _doc.ksArcBy3Points(-(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 2) * Math.Sin(DegToRad(0.375 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 2) * Math.Cos(DegToRad(0.375 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(alfa2)),
                            1);
                        _doc.ksDeleteMtr();
                        iSketch3Def.EndEdit();
                    }
                }
                // ������ ��������� ���������
                ksEntity iOffsetPlane1Entity = iPart.NewEntity((short) Obj3dType.o3d_planeOffset) as ksEntity;
                if (iOffsetPlane1Entity != null)
                {
                    ksPlaneOffsetDefinition iOffsetPlane1Def =
                        iOffsetPlane1Entity.GetDefinition() as ksPlaneOffsetDefinition;
                    if (iOffsetPlane1Def != null)
                    {
                        // �������� �������� �� ��
                        iOffsetPlane1Def.offset = thickness / 2;
                        // ����������� ���������������
                        iOffsetPlane1Def.direction = true;
                        iOffsetPlane1Def.SetPlane(planeYoz);
                        // ������ ��������� �������
                        iOffsetPlane1Entity.hidden = true;
                        // ������� ��������� ���������
                        iOffsetPlane1Entity.Create();
                    }
                }
                // ������ (���������) ����� ������ ����� �������
                ksEntity iSketch4Entity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketch4Entity != null)
                {
                    ksSketchDefinition iSketch4Def = iSketch4Entity.GetDefinition() as ksSketchDefinition;
                    if (iSketch4Def != null)
                    {
                        // ������� ��������� � ������ ��� ��������� ���������
                        iSketch4Def.SetPlane(iOffsetPlane1Entity);
                        iSketch4Entity.Create();
                        _doc = iSketch4Def.BeginEdit() as ksDocument2D;
                        alfa2 = -(180 / Math.PI * (2 * thickness * Math.Tan(Math.PI * (angle) / 180) / diameterPitch));
                        _doc.ksMtr(0, 0, 90, 1, 1);
                        // ������������ ����������� ������
                        // ������ ��������� ��� ��������
                        // ����� ������� ���� �� ���� ������
                        _doc.ksArcBy3Points(-(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(alfa2)),
                            -diameterPitch / 2 * Math.Sin(DegToRad(alfa1 / 8 + alfa2)),
                            -diameterPitch / 2 * Math.Cos(DegToRad(alfa1 / 8 + alfa2)),
                            -diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 4 + alfa2)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 4 + alfa2)), 1);
                        _doc.ksArcByPoint(0, 0, diameterIn / 2, -diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 4 + alfa2)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 4 + alfa2)),
                            -diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 2 + alfa2)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 2 + alfa2)), -1, 1);
                        _doc.ksArcBy3Points(-diameterIn / 2 * Math.Sin(DegToRad(alfa1 / 2 + alfa2)),
                            -diameterIn / 2 * Math.Cos(DegToRad(alfa1 / 2 + alfa2)),
                            -diameterPitch / 2 * Math.Sin(DegToRad(0.625 * alfa1 + alfa2)),
                            -diameterPitch / 2 * Math.Cos(DegToRad(0.625 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1 + alfa2)), 1);
                        _doc.ksArcBy3Points(-(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(0.75 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(0.75 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 2) * Math.Sin(DegToRad(0.375 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 2) * Math.Cos(DegToRad(0.375 * alfa1 + alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Sin(DegToRad(alfa2)),
                            -(diameterOut / 2 + 0.1) * Math.Cos(DegToRad(alfa2)),
                            1);
                        _doc.ksDeleteMtr();
                        iSketch4Def.EndEdit();
                    }
                }
                // ��������� �������� �������� �� ��������
                ksEntity iCutLoftEntity = iPart.NewEntity((short) Obj3dType.o3d_cutLoft) as ksEntity;
                if (iCutLoftEntity != null)
                {
                    // ��������� ���������� �������� �� ��������
                    ksCutLoftDefinition iCutLoftDef = iCutLoftEntity.GetDefinition() as ksCutLoftDefinition;
                    if (iCutLoftDef != null)
                    {
                        // ��������� ������� ksEntityCollection
                        // ��������� ������� ��� ��������� �� ��������
                        ksEntityCollection collect = iCutLoftDef.Sketchs() as ksEntityCollection;
                        // ��������� ������ � ��������
                        collect.Add(iSketch2Entity);
                        collect.Add(iSketch3Entity);
                        collect.Add(iSketch4Entity);
                        // ������� �������� �� ��������
                        // ��������� � ������ ����� ����� ������� � ����� ������
                        iCutLoftEntity.Create();
                    }
                }
                // ��������� ��������������� ��� �� ����������� ���� ����������
                ksEntity iAxis = iPart.NewEntity((short) Obj3dType.o3d_axis2Planes) as ksEntity;
                if (iAxis != null)
                {
                    // ��������� ���������� ��������������� ���
                    // �� ����������� ����������
                    ksAxis2PlanesDefinition iAxis2PlDef = iAxis.GetDefinition() as ksAxis2PlanesDefinition;
                    if (iAxis2PlDef != null)
                    {
                        // ������ ���������
                        iAxis2PlDef.SetPlane(1, planeXoz);
                        iAxis2PlDef.SetPlane(2, planeXoy);
                        // ������ ��� ���������
                        iAxis.hidden = true;
                        // ������� ��������������� ���
                        iAxis.Create();
                    }
                }
                // ��������� �������� ������ �� ��������������� �����
                ksEntity iCircularCopy = iPart.NewEntity((short) Obj3dType.o3d_circularCopy) as ksEntity;
                if (iCircularCopy != null)
                {
                    // ��������� ���������� �������� ����������� �� �������
                    ksCircularCopyDefinition iCirCopyDef = iCircularCopy.GetDefinition() as ksCircularCopyDefinition;
                    if (iCirCopyDef != null)
                    {
                        // ��������� �������� ��� �����������
                        ksEntityCollection collect1 = iCirCopyDef.GetOperationArray() as ksEntityCollection;
                        // �������� ����� ���� ���� � ��������� ����
                        collect1.Add(iCutLoftEntity);
                        // ���������� �����, ����� ���������� ������
                        iCirCopyDef.count2 = teethCount;
                        iCirCopyDef.factor2 = true;
                        // ��� �����������
                        iCirCopyDef.SetAxis(iAxis);
                        // ������� ��������������� ������ � ������ ������!
                        iCircularCopy.Create();
                    }
                }
            }
        }

        #region COM Registration

        /// <summary>
        /// ��� ������� ����������� ��� ����������� ������ ��� COM
        /// ��� ��������� � ����� ������� ���������� ������ Kompas_Library,
        /// ������� ������������� � ���, ��� ����� �������� ����������� ������,
        /// � ����� �������� ��� InprocServer32 �� ������, � ��������� ����.
        /// ��� ��� �������� ��� ����, ����� ����� ����������� ����������
        /// ���������� �� ������� ActiveX.
        /// </summary>
        [ComRegisterFunction]
        public static void RegisterKompasLib(Type t)
        {
            try
            {
                RegistryKey regKey = Registry.LocalMachine;
                string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
                regKey = regKey.OpenSubKey(keyName, true);
                regKey.CreateSubKey("Kompas_Library");
                regKey = regKey.OpenSubKey("InprocServer32", true);
                regKey.SetValue(null,
                    System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
                regKey.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("��� ����������� ������ ��� COM-Interop ��������� ������:\n{0}", ex));
            }
        }


        /// <summary>
        /// ��� ������� ������� ������ Kompas_Library �� �������
        /// </summary>
        /// <param name="t">The t.</param>
        [ComUnregisterFunction]
        public static void UnregisterKompasLib(Type t)
        {
            RegistryKey regKey = Registry.LocalMachine;
            string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
            RegistryKey subKey = regKey.OpenSubKey(keyName, true);
            subKey.DeleteSubKey("Kompas_Library");
            subKey.Close();
        }

        #endregion
    }
}



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
    /// Основной класс плагина. 
    /// Здесь находится логика создания детали и точка входа из Компаса.
    /// </summary>
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class GearPlugin
    {
        /// <summary>
        /// Интерфейс API КОМПАС
        /// </summary>
        private KompasObject _kompas;

        /// <summary>
        /// Интерфейс Automation документа-модели
        /// </summary>
        private ksDocument3D _doc3;

        /// <summary>
        /// Интерфейс графического документа системы КОМПАС
        /// </summary>
        private ksDocument2D _doc;


        // Имя библиотеки
        /// <summary>
        /// Получение имени библиотеки.
        /// </summary>
        /// <returns>Имя библиотеки</returns>
        [return: MarshalAs(UnmanagedType.BStr)]
        public string GetLibraryName()
        {
            return "Плагин";
        }

        /// <summary>
        /// Головная функция библиотеки
        /// </summary>
        /// <param name="command">Номер команды в меню</param>
        /// <param name="mode">Режим работы</param>
        /// <param name="kompas">интерфейс KompasObject</param>
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

                    _doc3.comment = "Зубчатое колесо";
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
        /// Получение строки меню для создания меню в виде строк
        /// </summary>
        /// <param name="number">Счетчик строк</param>
        /// <param name="itemType">Тип строки</param>
        /// <param name="command">Номер команды</param>
        /// <returns>Строка меню</returns>
        [return: MarshalAs(UnmanagedType.BStr)]
        public string ExternalMenuItem(short number, ref short itemType, ref short command)
        {
            string result = string.Empty; //По уполчанию - пустая строка
            itemType = 1; //MENUITEM

            switch (number)
            {
                case 1:
                    result = "Создать шестерню";
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
        /// Перевод градусов в радианы
        /// </summary>
        /// <param name="angle">Угол в градусах</param>
        /// <returns>Угол в радианах</returns>
        private static double DegToRad(double angle)
        {
            return Math.PI * angle / 180.0;
        }

        /// <summary>
        /// Создание модели шестерни
        /// </summary>
        /// <param name="properties">Передаваемые с формы параметры</param>
        public void CreateModel(Gear properties)
        {
            int teethCount = properties.TeethCount;
            var angle = properties.Angle;
            var thickness = properties.Thickness;
            var shaftDiam = properties.ShaftDiam;
            var diameterOut = properties.DiameterOut;
            var diameterIn = properties.DiameterIn;
            var diameterPitch = (diameterOut + diameterIn) / 2; // делительный диаметр колеса
            var alfa1 = 0.0;
            var alfa2 = 0.0;
            // интерфейсы ортогональных плоскостей
            ksPart iPart = _doc3.GetPart((int) Part_Type.pNew_Part) as ksPart; //новый компонет
            if (iPart != null)
            {
                // интерфейсы ортогональных плоскостей
                ksEntity planeXoy = iPart.GetDefaultEntity((short) Obj3dType.o3d_planeXOY) as ksEntity;
                ksEntity planeXoz = iPart.GetDefaultEntity((short) Obj3dType.o3d_planeXOZ) as ksEntity;
                ksEntity planeYoz = iPart.GetDefaultEntity((short) Obj3dType.o3d_planeYOZ) as ksEntity;
                // интерфейс эскиза (половина контура сечения колеса)
                ksEntity iSketchEntity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketchEntity != null)
                {
                    // интерфейс параметров эскиза
                    ksSketchDefinition iSketchDef = iSketchEntity.GetDefinition() as ksSketchDefinition;
                    if (iSketchDef != null)
                    {
                        if (planeXoy != null)
                        {
                            // устанавливаем плоскость,
                            // на которой создается эскиз
                            iSketchDef.SetPlane(planeXoy);

                            iSketchEntity.Create();

                            // запускаем процесс редактирования эскиза
                            // doc – указатель на интерфейс ksDocument2D
                            _doc = iSketchDef.BeginEdit() as ksDocument2D;
                            if (_doc != null)
                            {
                                _doc.ksLineSeg(-thickness / 2, 00, thickness / 2, 0, 3); //3 - осевая линия
                                _doc.ksLineSeg(thickness / 2, shaftDiam / 2, -thickness / 2, shaftDiam / 2, 1);
                                _doc.ksLineSeg(thickness / 2, shaftDiam / 2, thickness / 2, diameterOut / 2, 1);
                                _doc.ksLineSeg(-thickness / 2, diameterOut / 2, -thickness / 2, shaftDiam / 2, 1);
                                _doc.ksLineSeg(-thickness / 2, diameterOut / 2, thickness / 2, diameterOut / 2, 1);
                            }
                            iSketchDef.EndEdit();
                        }
                    }
                }
                // интерфейс базовой операции вращения
                ksEntity iBaseRotatedEntity = iPart.NewEntity((short) Obj3dType.o3d_baseRotated) as ksEntity;
                if (iBaseRotatedEntity != null)
                {
                    // интерфейс параметров вращения
                    ksBaseRotatedDefinition iBaseRotatedDef =
                        iBaseRotatedEntity.GetDefinition() as ksBaseRotatedDefinition;
                    if (iBaseRotatedDef != null)
                    {
                        // настройка параметров вращения
                        iBaseRotatedDef.SetThinParam(false, (short) Direction_Type.dtNormal, 1, 1);
                        iBaseRotatedDef.SetSideParam(true, 360.0);
                        iBaseRotatedDef.toroidShapeType = false;
                        iBaseRotatedDef.SetSketch(iSketchEntity);
                        // создаем операцию вращения
                        // результат – заготовка зубчатого колеса
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
                            // размещаем эскиз на плоскости XOY
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
                // интерфейс операции Вырезать выдавливанием
                ksEntity iCutExtrusion = iPart.NewEntity((short) Obj3dType.o3d_cutExtrusion) as ksEntity;
                if (iCutExtrusion != null)
                {
                    // интерфейс параметров вырезания
                    ksCutExtrusionDefinition iCutExtrusionDef =
                        iCutExtrusion.GetDefinition() as ksCutExtrusionDefinition;
                    if (iCutExtrusionDef != null)
                    {
                        // настройка параметров 
                        iCutExtrusionDef.SetSketch(iSketch1Entity);
                        // направление
                        iCutExtrusionDef.directionType = (short) Direction_Type.dtNormal;
                        // величина вырезания по каждому из направлений
                        iCutExtrusionDef.SetSideParam(true, (short) ksEndTypeEnum.etBlind,
                            properties.KeywayDepth + shaftDiam / 2, 0, false);
                        iCutExtrusionDef.SetThinParam(false, 0, 0, 0);
                        iCutExtrusion.Create();
                    }
                }
                // интерфейс смещенной плоскости
                ksEntity iOffsetPlaneEntity = iPart.NewEntity((short) Obj3dType.o3d_planeOffset) as ksEntity;
                if (iOffsetPlaneEntity != null)
                {
                    // интерфейс параметров смещенной плоскости
                    ksPlaneOffsetDefinition iOffsetPlaneDef =
                        iOffsetPlaneEntity.GetDefinition() as ksPlaneOffsetDefinition;
                    if (iOffsetPlaneDef != null)
                    {
                        // величина, базовая плоскость и другие параметры смещения
                        iOffsetPlaneDef.offset = thickness / 2;
                        iOffsetPlaneDef.SetPlane(planeYoz);
                        iOffsetPlaneDef.direction = false;
                        // делаем плоскость скрытой
                        iOffsetPlaneEntity.hidden = true;
                        // создаем вспомогательную плоскость
                        iOffsetPlaneEntity.Create();
                    }
                }
                // эскиз первого выреза между зубьями
                ksEntity iSketch2Entity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketch2Entity != null)
                {
                    ksSketchDefinition iSketch2Def = iSketch2Entity.GetDefinition() as ksSketchDefinition;
                    if (iSketch2Def != null)
                    {
                        // базовая плоскость – вспомогательная iOffsetPlaneEntity
                        iSketch2Def.SetPlane(iOffsetPlaneEntity);
                        iSketch2Entity.Create();
                        _doc = iSketch2Def.BeginEdit() as ksDocument2D;
                        alfa1 = 360.0 / teethCount;
                        _doc.ksMtr(0, 0, 90, 1, 1);
                        // вычерчивание изображения эскиза
                        // вместо эвольвент для простоты
                        // берем обычные дуги по трем точкам
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
                // интерфейс второго эскиза выреза между зубьями
                ksEntity iSketch3Entity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketch3Entity != null)
                {
                    ksSketchDefinition iSketch3Def = iSketch3Entity.GetDefinition() as ksSketchDefinition;
                    if (iSketch3Def != null)
                    {
                        // строим на плоскости YOZ
                        iSketch3Def.SetPlane(planeYoz);
                        iSketch3Entity.Create();
                        _doc = iSketch3Def.BeginEdit() as ksDocument2D;
                        alfa2 = -(180 / Math.PI * (thickness * Math.Tan(Math.PI * (angle) / 180) / diameterPitch));
                        _doc.ksMtr(0, 0, 90, 1, 1);
                        // вычерчивание изображения эскиза
                        // вместо эвольвент для простоты
                        // берем обычные дуги по трем точкам
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
                // вторая смещенная плоскость
                ksEntity iOffsetPlane1Entity = iPart.NewEntity((short) Obj3dType.o3d_planeOffset) as ksEntity;
                if (iOffsetPlane1Entity != null)
                {
                    ksPlaneOffsetDefinition iOffsetPlane1Def =
                        iOffsetPlane1Entity.GetDefinition() as ksPlaneOffsetDefinition;
                    if (iOffsetPlane1Def != null)
                    {
                        // величина смещения та же
                        iOffsetPlane1Def.offset = thickness / 2;
                        // направление противоположное
                        iOffsetPlane1Def.direction = true;
                        iOffsetPlane1Def.SetPlane(planeYoz);
                        // делаем плоскость скрытой
                        iOffsetPlane1Entity.hidden = true;
                        // создаем смещенную плоскость
                        iOffsetPlane1Entity.Create();
                    }
                }
                // третий (последний) эскиз выреза между зубьями
                ksEntity iSketch4Entity = iPart.NewEntity((short) Obj3dType.o3d_sketch) as ksEntity;
                if (iSketch4Entity != null)
                {
                    ksSketchDefinition iSketch4Def = iSketch4Entity.GetDefinition() as ksSketchDefinition;
                    if (iSketch4Def != null)
                    {
                        // базовая плоскость – только что созданная смещенная
                        iSketch4Def.SetPlane(iOffsetPlane1Entity);
                        iSketch4Entity.Create();
                        _doc = iSketch4Def.BeginEdit() as ksDocument2D;
                        alfa2 = -(180 / Math.PI * (2 * thickness * Math.Tan(Math.PI * (angle) / 180) / diameterPitch));
                        _doc.ksMtr(0, 0, 90, 1, 1);
                        // вычерчивание изображения эскиза
                        // вместо эвольвент для простоты
                        // берем обычные дуги по трем точкам
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
                // интерфейс операции Вырезать по сечениям
                ksEntity iCutLoftEntity = iPart.NewEntity((short) Obj3dType.o3d_cutLoft) as ksEntity;
                if (iCutLoftEntity != null)
                {
                    // интерфейс параметров операции по сечениям
                    ksCutLoftDefinition iCutLoftDef = iCutLoftEntity.GetDefinition() as ksCutLoftDefinition;
                    if (iCutLoftDef != null)
                    {
                        // интерфейс массива ksEntityCollection
                        // коллекции эскизов для вырезания по сечениям
                        ksEntityCollection collect = iCutLoftDef.Sketchs() as ksEntityCollection;
                        // добавляем эскизы в колекцию
                        collect.Add(iSketch2Entity);
                        collect.Add(iSketch3Entity);
                        collect.Add(iSketch4Entity);
                        // создаем операцию по сечениям
                        // результат – первый вырез между зубьями в венце колеса
                        iCutLoftEntity.Create();
                    }
                }
                // интерфейс вспомогательной оси на пересечении двух плоскостей
                ksEntity iAxis = iPart.NewEntity((short) Obj3dType.o3d_axis2Planes) as ksEntity;
                if (iAxis != null)
                {
                    // интерфейс параметров вспомогательной оси
                    // на пересечении плоскостей
                    ksAxis2PlanesDefinition iAxis2PlDef = iAxis.GetDefinition() as ksAxis2PlanesDefinition;
                    if (iAxis2PlDef != null)
                    {
                        // задаем плоскости
                        iAxis2PlDef.SetPlane(1, planeXoz);
                        iAxis2PlDef.SetPlane(2, planeXoy);
                        // делаем ось невидимой
                        iAxis.hidden = true;
                        // создаем вспомогательную ось
                        iAxis.Create();
                    }
                }
                // интерфейс операции Массив по концентрической сетке
                ksEntity iCircularCopy = iPart.NewEntity((short) Obj3dType.o3d_circularCopy) as ksEntity;
                if (iCircularCopy != null)
                {
                    // интерфейс параметров операции копирования по массиву
                    ksCircularCopyDefinition iCirCopyDef = iCircularCopy.GetDefinition() as ksCircularCopyDefinition;
                    if (iCirCopyDef != null)
                    {
                        // коллекция операций для копирования
                        ksEntityCollection collect1 = iCirCopyDef.GetOperationArray() as ksEntityCollection;
                        // операция всего лишь одна – вырезание зуба
                        collect1.Add(iCutLoftEntity);
                        // количество копий, равно количеству зубьев
                        iCirCopyDef.count2 = teethCount;
                        iCirCopyDef.factor2 = true;
                        // ось копирования
                        iCirCopyDef.SetAxis(iAxis);
                        // создаем концентрический массив – колесо готово!
                        iCircularCopy.Create();
                    }
                }
            }
        }

        #region COM Registration

        /// <summary>
        /// Эта функция выполняется при регистрации класса для COM
        /// Она добавляет в ветку реестра компонента раздел Kompas_Library,
        /// который сигнализирует о том, что класс является приложением Компас,
        /// а также заменяет имя InprocServer32 на полное, с указанием пути.
        /// Все это делается для того, чтобы иметь возможность подключить
        /// библиотеку на вкладке ActiveX.
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
                MessageBox.Show(string.Format("При регистрации класса для COM-Interop произошла ошибка:\n{0}", ex));
            }
        }


        /// <summary>
        /// Эта функция удаляет раздел Kompas_Library из реестра
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

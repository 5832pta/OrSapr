namespace Steps.NET
{
    /// <summary>
    /// Параметры шестерни
    /// </summary>
    public class Gear
    {
        /// <summary>
        /// Диаметр окружности впадин
        /// </summary>
        public double DiameterIn { get; set; }

        /// <summary>
        /// Диаметр окружности вершин
        /// </summary>
        public double DiameterOut { get; set; }

        /// <summary>
        /// Диаметр отверстия
        /// </summary>
        public double ShaftDiam { get; set; }

        /// <summary>
        /// Толщина детали
        /// </summary>
        public double Thickness { get; set; }

        /// <summary>
        /// Глубина шпоночного паза
        /// </summary>
        public double KeywayDepth { get; set; }

        /// <summary>
        /// Ширина шпоночного паза
        /// </summary>
        public double KeywayWidth { get; set; }

        /// <summary>
        /// Угол наклона зубьев
        /// </summary>
        public double Angle { get; set; }

        /// <summary>
        /// Количество зубьев
        /// </summary>
        public short TeethCount { get; set; }
    }
}

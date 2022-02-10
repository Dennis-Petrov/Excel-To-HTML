namespace Excel_To_HTML
{
    /// <summary>
    /// DTO для формирования подписей в синих штампах.
    /// </summary>
    public class Sign
    {
        /// <summary>
        /// Инициализирует экземпляр <see cref="Sign"/> класса.
        /// </summary>
        /// <param name="certificate"> Сертификат подписи для создания синего штампа. </param>
        /// <param name="signingTime"> Дата подписания для создания синего штампа. </param>
        /// <param name="signatureTimeStampTime"> Штамп времени для создания синего штампа. </param>
        public Sign(PrintformCertificate certificate, DateTime? signingTime, DateTime? signatureTimeStampTime)
        {
            OrganizationName = certificate.OrganizationName;
            Employee = $"{certificate.Surname} {certificate.FirstName} {certificate.LastName}";
            SerialNumber = certificate.SerialNumber;
            ValidityPeriod = $"{certificate.ValidFrom:dd.MM.yyyy HH:mm:ss} - {certificate.ValidTo:dd.MM.yyyy HH:mm:ss} GMT +3";
            SigningTime = signingTime != null ? $"{signingTime.Value:dd.MM.yyyy HH:mm:ss} GMT +3" : string.Empty;
            SignatureTimeStampTime = signatureTimeStampTime != null ? $"{signatureTimeStampTime.Value:dd.MM.yyyy HH:mm:ss} GMT +3" : string.Empty;
        }

        /// <summary>
        /// Задает или возвращает время подписания документа.
        /// </summary>
        public string SigningTime { get; set; }

        /// <summary>
        /// Задает или возвращает время заверения подписи документа.
        /// </summary>
        public string SignatureTimeStampTime { get; set; }

        /// <summary>
        /// Задает или возвращает наименование подписанта организации.
        /// </summary>
        public string OrganizationName { get; set; }

        /// <summary>
        /// Задает или возвращает имя ответственного лица.
        /// </summary>
        public string Employee { get; set; }

        /// <summary>
        /// Задает или возвращает имя серийный номер сертификата.
        /// </summary>
        public string SerialNumber { get; set; }

        /// <summary>
        /// Задает или возвращает имя период действия сертификата.
        /// </summary>
        public string ValidityPeriod { get; set; }
    }
}

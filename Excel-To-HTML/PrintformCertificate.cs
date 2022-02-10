namespace Excel_To_HTML
{
    /// <summary>
    /// Определяет сертификат, который был использован для формирования подписи.
    /// </summary>
    public sealed class PrintformCertificate
    {
        /// <summary>
        /// Возвращает или задает имя.
        /// </summary>
        public string? FirstName { get; set; }

        /// <summary>
        /// Возвращает или задает фамилию.
        /// </summary>
        public string? Surname { get; set; }

        /// <summary>
        /// Возвращает или задает отчество.
        /// </summary>
        public string? LastName { get; set; }

        /// <summary>
        /// Возвращает или задает наименование организации.
        /// </summary>
        public string? OrganizationName { get; set; }

        /// <summary>
        /// Возвращает или задает серийный номер.
        /// </summary>
        public string? SerialNumber { get; set; }

        /// <summary>
        /// Возвращает или задает дату, начиная с которой сертификат валиден.
        /// </summary>
        public DateTime ValidFrom { get; set; }

        /// <summary>
        /// Возвращает или задает дату, по которую сертификат валиден.
        /// </summary>
        public DateTime ValidTo { get; set; }
    }
}

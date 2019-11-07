using System;

namespace DocumentReview
{
    public class Author
    {
        private string _initials;

        public Author(string name, string initials)
        {
            Name = name;
            Initials = initials;
        }

        public string Name { get; set; }

        public string Initials
        {
            get => _initials;
            set
            {
                if (value.Length != 3)
                {
                    throw new ArgumentException("Инициалы должны быть из 3х букв.");
                }

                _initials = value;
            }
        }
    }
}
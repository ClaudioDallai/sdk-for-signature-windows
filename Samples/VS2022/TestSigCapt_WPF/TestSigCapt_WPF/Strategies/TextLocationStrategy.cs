using iTextSharp.text.pdf.parser;
using System.Collections.Generic;
using System.Linq;


namespace TestSigCapt_WPF
{
    public class TextLocationStrategyChar : IRenderListener
    {
        private readonly string _target;
        private readonly List<(string text, Vector pos)> _chars = new List<(string, Vector)>();
        public List<Vector> MarkerPositions { get; } = new List<Vector>();

        public TextLocationStrategyChar(string target)
        {
            _target = target;
        }

        public void RenderText(TextRenderInfo renderInfo)
        {
            foreach (var chrInfo in renderInfo.GetCharacterRenderInfos())
            {
                _chars.Add((chrInfo.GetText(), chrInfo.GetBaseline().GetStartPoint()));
            }
        }

        public void BeginTextBlock() { }
        public void EndTextBlock() { }
        public void RenderImage(ImageRenderInfo renderInfo) { }

        public void FinalizeSearch()
        {
            string fullText = string.Concat(_chars.Select(c => c.text));
            int index = 0;

            while ((index = fullText.IndexOf(_target, index)) >= 0)
            {
                MarkerPositions.Add(_chars[index].pos);
                index += _target.Length;
            }
        }
    }

}

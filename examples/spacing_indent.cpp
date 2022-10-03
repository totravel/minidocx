
#include "minidocx.hpp"

#define text1 "I feel that there is much to be said for the Celtic belief that the souls of those whom we have lost are held captive in some inferior being, in an animal, in a plant, in some inanimate object, and so effectively lost to us until the day (which to many never comes) when we happen to pass by the tree or to obtain possession of the object which forms their prison."
#define text2 "Then they start and tremble, they call us by our name, and as soon as we have recognized their voice the spell is broken. We have delivered them: they have overcome death and return to share our life. And so it is with our own past. It is a labor in vain to attempt to recapture it: all the efforts of our intellect must prove futile."
#define text3 "The past is hidden somewhere outside the realm, beyond the reach of intellect, in some material object (in the sensation which that material object will give us) which we do not suspect. And as for that object, it depends on chance whether we come upon it or not before we ourselves must die."

int main()
{
  docx::Document doc("./spacing_indent.docx");

  // Page 1
  auto p1 = doc.AppendParagraph(text1);
  auto p2 = doc.AppendParagraph(text2);
  auto p3 = doc.AppendParagraph(text3);
  auto p4 = doc.AppendParagraph(text1);
  auto p5 = doc.AppendParagraph(text2);
  auto p6 = doc.AppendParagraph(text3);

  p2.SetFontStyle(docx::Run::Bold);
  p4.SetFontStyle(docx::Run::Bold);
  p6.SetFontStyle(docx::Run::Bold);

  // Line spacing
  p1.SetLineSpacingSingle();                   // Single
  p2.SetLineSpacingLines(1.5);                 // 1.5 lines
  p3.SetLineSpacingLines(2);                   // Double (2 lines)
  p4.SetLineSpacingAtLeast(docx::Pt2Twip(12)); // At Least (12 pt)
  p5.SetLineSpacingExactly(docx::Pt2Twip(12)); // Exactly (12 pt)
  p6.SetLineSpacingLines(3);                   // Multiple (3 lines)

  // Indent
  p1.SetLeftIndentChars(2);
  p2.SetRightIndent(docx::CM2Twip(3));
  p3.SetFirstLineChars(2);
  p4.SetFirstLine(docx::CM2Twip(2));
  p5.SetHangingChars(2);
  p6.SetHanging(docx::CM2Twip(2));

  // Page 2
  doc.AppendPageBreak();
  auto p7 = doc.AppendParagraph("This is the 7th paragraph.");
  auto p8 = doc.AppendParagraph("This is the 8th paragraph.");
  auto p9 = doc.AppendParagraph("This is the 9th paragraph.");
  auto p10 = doc.AppendParagraph("This is the 10th paragraph.");

  // Spacing
  p8.SetBeforeSpacingAuto();
  p9.SetBeforeSpacingLines(1.5);
  p9.SetAfterSpacing(docx::Pt2Twip(10));

  doc.Save();
  return 0;
}

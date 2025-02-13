
#pragma once

namespace NAMESPACE
{
  //   mm    cm    in    pt    tw    emu
  //    1                          36000
  //          1                   360000
  // 25.4  2.54     1    72  1440 914400
  //                      1    20
  //                            1    635

  // Page sizes in tw.

  // A3
  const unsigned int A3_W = 16838;
  const unsigned int A3_H = 23811;

  // A4
  const unsigned int A4_W = 11906; // 210 mm * 36000 / 635
  const unsigned int A4_H = 16838; // 297 mm * 36000 / 635

  // Letter
  const unsigned int LETTER_W = 12240;
  const unsigned int LETTER_H = 15840;

  // Legal
  const unsigned int LEGAL_W = 12240;
  const unsigned int LEGAL_H = 20160;

  // Tabloid
  const unsigned int TABLOID_W = 15840;
  const unsigned int TABLOID_H = 24480;

  // Executive
  const unsigned int EXECUTIVE_W = 10440;
  const unsigned int EXECUTIVE_H = 15120;

}

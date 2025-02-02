
#pragma once

namespace NAMESPACE
{
  const double A0_W = 33.1; // page size in inches
  const double A0_H = 46.8;

  const double A1_W = 23.4;
  const double A1_H = 33.1;

  const double A2_W = 16.5;
  const double A2_H = 23.4;

  const double A3_W = 11.7;
  const double A3_H = 16.5;

  const double A4_W = 8.3;
  const double A4_H = 11.7;

  const double A5_W = 5.8;
  const double A5_H = 8.3;

  const double A6_W = 4.1;
  const double A6_H = 5.8;

  const double LETTER_W = 8.5;
  const double LETTER_H = 11;

  const double LEGAL_W = 8.5;
  const double LEGAL_H = 14;

  const double TABLOID_W = 11;
  const double TABLOID_H = 17;


  const unsigned int A0_COLS = 2384; // page size in pixels (PPI = 72)
  const unsigned int A0_ROWS = 3370;

  const unsigned int A1_COLS = 1684;
  const unsigned int A1_ROWS = 2384;

  const unsigned int A2_COLS = 1191;
  const unsigned int A2_ROWS = 1684;

  const unsigned int A3_COLS = 842;
  const unsigned int A3_ROWS = 1190;

  const unsigned int A4_COLS = 595;
  const unsigned int A4_ROWS = 842;

  const unsigned int A5_COLS = 420;
  const unsigned int A5_ROWS = 595;

  const unsigned int A6_COLS = 297;
  const unsigned int A6_ROWS = 420;

  const unsigned int LETTER_COLS = 612;
  const unsigned int LETTER_ROWS = 792;

  const unsigned int LEGAL_COLS = 612;
  const unsigned int LEGAL_ROWS = 1008;

  const unsigned int TABLOID_COLS = 792;
  const unsigned int TABLOID_ROWS = 1224;


  inline long long pt2twip(const double pt) // 1 pt = 20 twip
  {
    return pt * 20;
  }

  inline double twip2pt(const long long twip)
  {
    return twip / 20.0;
  }

  inline double inch2pt(const double inch) // 1 inch = 72 pt
  {
    return inch * 72;
  }

  inline double pt2inch(const double pt)
  {
    return pt / 72;
  }

  inline double mm2inch(const long long mm) // 25.4 mm = 1 inch
  {
    return mm / 25.4;
  }

  inline long long inch2mm(const double inch)
  {
    return inch * 25.4;
  }

  inline double cm2inch(const double cm) // 2.54 cm = 1 inch
  {
    return cm / 2.54;
  }

  inline double inch2cm(const double inch)
  {
    return inch * 2.54;
  }

  inline long long inch2twip(const double inch) // 1 inch = 72 pt = 1440 twip
  {
    return inch * 1440;
  }

  inline double twip2inch(const long long twip)
  {
    return twip / 1440.0;
  }

  inline long long inch2emu(const double inch) // 1 inch = 914400 EMUs (English Metric Units)
  {
    return inch * 914400;
  }

  inline double emu2inch(const long long emu)
  {
    return emu / 914400;
  }

  inline long long cm2emu(const double cm) // 1 cm = 360000 EMUs
  {
    return cm * 360000;
  }

  inline double emu2cm(const long long emu)
  {
    return emu / 360000;
  }

  inline long long mm2twip(const long long mm)
  {
    return inch2twip(mm2inch(mm));
  }

  inline long long twip2mm(const long long twip)
  {
    return inch2mm(twip2inch(twip));
  }

  inline long long cm2twip(const double cm)
  {
    return inch2twip(cm2inch(cm));
  }

  inline double twip2cm(const long long twip)
  {
    return inch2cm(twip2inch(twip));
  }
}

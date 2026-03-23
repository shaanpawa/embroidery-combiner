"use client";

import { createContext, useContext, useState, useEffect, ReactNode } from "react";

type Lang = "en" | "th";

interface I18nContext {
  lang: Lang;
  toggle: () => void;
  t: (key: string) => string;
}

const translations: Record<string, Record<Lang, string>> = {
  // Login
  "login.welcome": { en: "Welcome back", th: "ยินดีต้อนรับ" },
  "login.subtitle": { en: "Sign in to access Micro Automation", th: "เข้าสู่ระบบ Micro Automation" },
  "login.password": { en: "Enter password", th: "ใส่รหัสผ่าน" },
  "login.signin": { en: "Sign In", th: "เข้าสู่ระบบ" },
  "login.signing_in": { en: "Signing in...", th: "กำลังเข้าสู่ระบบ..." },
  "login.google": { en: "or sign in with Google", th: "หรือเข้าสู่ระบบด้วย Google" },
  "login.error.password": { en: "Incorrect password. Please try again.", th: "รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่" },
  "login.error.denied": { en: "Access denied. Contact your administrator.", th: "ไม่มีสิทธิ์เข้าใช้งาน ติดต่อผู้ดูแลระบบ" },
  "login.error.generic": { en: "Something went wrong. Please try again.", th: "เกิดข้อผิดพลาด กรุณาลองใหม่" },
  "login.footer": { en: "MICRO AUTOMATION BY OSSIA", th: "MICRO AUTOMATION BY OSSIA" },

  // Homepage
  "home.subtitle": { en: "AUTOMATION TOOLS", th: "เครื่องมืออัตโนมัติ" },
  "home.combo.title": { en: "Embroidery Stacker", th: "Embroidery Stacker" },
  "home.combo.desc": { en: "Stack embroidery name programs into production-ready combo files.", th: "รวมโปรแกรมชื่อปักเป็นไฟล์คอมโบสำหรับการผลิต" },
  "home.combo.available": { en: "AVAILABLE", th: "พร้อมใช้งาน" },
  "home.stitch.title": { en: "Stitch Count Predictor", th: "ตัวทำนายจำนวนฝีเข็ม" },
  "home.stitch.desc": { en: "Predict stitch counts and production time from design files.", th: "ทำนายจำนวนฝีเข็มและเวลาผลิตจากไฟล์ดีไซน์" },
  "home.batch.title": { en: "Batch Inspector", th: "ตรวจสอบแบทช์" },
  "home.batch.desc": { en: "Validate and inspect DST/NGS files before production runs.", th: "ตรวจสอบไฟล์ DST/NGS ก่อนผลิต" },
  "home.coming_soon": { en: "Coming Soon", th: "เร็วๆ นี้" },
  "home.footer": { en: "PRODUCTION TOOLS FOR MICRO EMBROIDERY CO.", th: "เครื่องมือการผลิตสำหรับ MICRO EMBROIDERY CO." },

  // Stacker - Session Picker
  "cb.title": { en: "Embroidery Stacker", th: "Embroidery Stacker" },
  "cb.session.subtitle": { en: "Create a new session to start stacking embroidery name programs into production files.", th: "สร้างเซสชันใหม่เพื่อเริ่มรวมโปรแกรมชื่อปักเป็นไฟล์การผลิต" },
  "cb.session.name": { en: "Session Name", th: "ชื่อเซสชัน" },
  "cb.session.start": { en: "Start Session", th: "เริ่มเซสชัน" },
  "cb.session.demo": { en: "or load demo data to explore", th: "หรือโหลดข้อมูลตัวอย่าง" },
  "cb.session.previous": { en: "Previous Sessions", th: "เซสชันก่อนหน้า" },
  "cb.session.loading": { en: "Loading session...", th: "กำลังโหลดเซสชัน..." },
  "cb.session.back": { en: "← Back to sessions", th: "← กลับไปเซสชัน" },
  "cb.session.new": { en: "+ New Session", th: "+ เซสชันใหม่" },

  // Stacker - Workflow
  "cb.subtitle": { en: "Stack embroidery name programs into production combo files", th: "รวมโปรแกรมชื่อปักเป็นไฟล์คอมโบสำหรับการผลิต" },
  "cb.step1": { en: "Step 1", th: "ขั้นตอน 1" },
  "cb.step2": { en: "Step 2", th: "ขั้นตอน 2" },
  "cb.step3": { en: "Step 3", th: "ขั้นตอน 3" },
  "cb.step.upload_order": { en: "Upload Order", th: "อัปโหลดคำสั่งซื้อ" },
  "cb.step.upload_programs": { en: "Upload Programs", th: "อัปโหลดโปรแกรม" },
  "cb.step.export": { en: "Export", th: "ส่งออก" },
  "cb.excel.title": { en: "Order Excel", th: "ไฟล์ Excel คำสั่งซื้อ" },
  "cb.excel.hint": { en: "Drop .xlsx or click to browse", th: "ลาก .xlsx หรือคลิกเพื่อเลือกไฟล์" },
  "cb.excel.parsing": { en: "Parsing...", th: "กำลังอ่าน..." },
  "cb.dst.title": { en: "DST Programs", th: "โปรแกรม DST" },
  "cb.dst.hint": { en: "Drop folder or click to browse .zip", th: "ลากโฟลเดอร์หรือคลิกเพื่อเลือก .zip" },
  "cb.dst.hint_disabled": { en: "Upload Excel first", th: "อัปโหลด Excel ก่อน" },
  "cb.dst.uploading": { en: "Uploading...", th: "กำลังอัปโหลด..." },
  "cb.dst.matched": { en: "programs matched", th: "โปรแกรมตรงกัน" },
  "cb.dst.uploaded": { en: "DST files uploaded", th: "ไฟล์ DST อัปโหลดแล้ว" },
  "cb.dst.missing": { en: "missing", th: "ไม่พบ" },

  // Column Mapping
  "cb.mapping.title": { en: "Column Mapping", th: "การจับคู่คอลัมน์" },
  "cb.mapping.auto": { en: "Auto-detected columns — verify and confirm", th: "ตรวจจับคอลัมน์อัตโนมัติ — ตรวจสอบและยืนยัน" },
  "cb.mapping.review": { en: "Some columns need manual selection", th: "บางคอลัมน์ต้องเลือกด้วยตนเอง" },
  "cb.mapping.needs_review": { en: "Needs review", th: "ต้องตรวจสอบ" },
  "cb.mapping.confirm": { en: "Confirm & Parse", th: "ยืนยันและอ่านข้อมูล" },
  "cb.mapping.select": { en: "— select —", th: "— เลือก —" },
  "cb.mapping.click_field": { en: "Click a field card to select it, then click a column to assign", th: "คลิกการ์ดฟิลด์ แล้วคลิกคอลัมน์เพื่อกำหนด" },
  "cb.mapping.click_column": { en: "Now click a column to assign", th: "คลิกคอลัมน์เพื่อกำหนด" },
  "cb.mapping.click_to_assign": { en: "click a column to assign", th: "คลิกคอลัมน์เพื่อกำหนด" },
  "cb.mapping.spreadsheet_preview": { en: "Spreadsheet Preview", th: "ตัวอย่างสเปรดชีต" },
  "cb.mapping.all_assigned": { en: "All required fields assigned", th: "กำหนดฟิลด์ที่จำเป็นครบแล้ว" },
  "cb.mapping.fields_remaining": { en: "required fields unassigned", th: "ฟิลด์ที่จำเป็นยังไม่ได้กำหนด" },
  "cb.mapping.field.program": { en: "Program", th: "โปรแกรม" },
  "cb.mapping.field.name_line1": { en: "Name", th: "ชื่อ" },
  "cb.mapping.field.name_line2": { en: "2nd Line", th: "บรรทัดที่ 2" },
  "cb.mapping.field.quantity": { en: "Quantity", th: "จำนวน" },
  "cb.mapping.field.com_no": { en: "Combo No", th: "เลขคอมโบ" },
  "cb.mapping.field.machine_program": { en: "MA", th: "MA" },
  "cb.mapping.help.program": { en: "DST file number (e.g., 42 = 42.DST)", th: "เลขไฟล์ DST (เช่น 42 = 42.DST)" },
  "cb.mapping.help.name_line1": { en: "Name to be embroidered", th: "ชื่อที่จะปัก" },
  "cb.mapping.help.name_line2": { en: "Second line — last name, title, organization (optional)", th: "บรรทัดที่ 2 — นามสกุล, ตำแหน่ง, องค์กร (ไม่บังคับ)" },
  "cb.mapping.help.quantity": { en: "How many copies of this name", th: "จำนวนสำเนาของชื่อนี้" },
  "cb.mapping.help.com_no": { en: "Names with same number go in the same output file", th: "ชื่อที่มีเลขเดียวกันจะรวมอยู่ในไฟล์เดียวกัน" },
  "cb.mapping.help.machine_program": { en: "MA code (e.g., MA50310)", th: "รหัส MA (เช่น MA50310)" },
  "cb.mapping.confidence_high": { en: "High confidence", th: "ความมั่นใจสูง" },
  "cb.mapping.confidence_medium": { en: "Review suggested", th: "แนะนำให้ตรวจสอบ" },
  "cb.mapping.confidence_low": { en: "Manual review required", th: "ต้องตรวจสอบด้วยตนเอง" },
  "cb.mapping.auto_detected": { en: "auto", th: "อัตโนมัติ" },
  "cb.mapping.position_guessed": { en: "guessed", th: "คาดเดา" },
  "cb.mapping.warn.qty_high": { en: "Quantities exceed 50 — verify", th: "จำนวนเกิน 50 — ตรวจสอบ" },
  "cb.mapping.warn.program_format": { en: "Should be numbers (DST file IDs)", th: "ควรเป็นตัวเลข (ID ไฟล์ DST)" },
  "cb.mapping.warn.combo_format": { en: "Contains letters — expected numbers", th: "มีตัวอักษร — ควรเป็นตัวเลข" },
  "cb.mapping.warn.ma_format": { en: "Doesn't match MA pattern (e.g., MA50310)", th: "ไม่ตรงรูปแบบ MA (เช่น MA50310)" },
  "cb.mapping.warn_low_title": { en: "⚠ Column detection failed", th: "⚠ ตรวจจับคอลัมน์ไม่สำเร็จ" },
  "cb.mapping.warn_low": { en: "Headers weren't recognised — every mapping below is a position-based guess and may be wrong. Verify each field carefully before confirming.", th: "ไม่พบส่วนหัวที่รู้จัก — การจับคู่ทั้งหมดเป็นการเดาจากตำแหน่ง กรุณาตรวจสอบทุกช่องอย่างละเอียดก่อนยืนยัน" },
  "cb.mapping.warn_medium_title": { en: "⚠ Some columns need review", th: "⚠ บางคอลัมน์ต้องตรวจสอบ" },
  "cb.mapping.warn_medium": { en: "Not all columns were confidently detected. Check each mapping below and correct any that look wrong.", th: "ไม่สามารถตรวจจับคอลัมน์ทั้งหมดได้อย่างมั่นใจ กรุณาตรวจสอบการจับคู่แต่ละช่อง" },
  "cb.mapping.how_it_works": { en: "How combining works", th: "วิธีการรวมไฟล์" },
  "cb.mapping.how_grouping": { en: "Rows with the same Combo No + MA are combined into one output file.", th: "แถวที่มี Combo No และ MA เดียวกันจะรวมอยู่ในไฟล์ผลลัพธ์เดียวกัน" },
  "cb.mapping.how_quantity": { en: "Quantity sets how many embroidery slots each row gets.", th: "Quantity กำหนดจำนวนช่องปักของแต่ละแถว" },
  "cb.mapping.how_program": { en: "Program maps to the DST filename (e.g., 42 → 42.DST).", th: "Program คือเลขไฟล์ DST (เช่น 42 → 42.DST)" },
  "cb.mapping.how_names": { en: "Name Line 1 & 2 are the text embroidered in each slot.", th: "ชื่อบรรทัด 1 และ 2 คือข้อความที่ปักในแต่ละช่อง" },

  // Stats
  "cb.stats.names": { en: "NAMES", th: "ชื่อ" },
  "cb.stats.groups": { en: "GROUPS", th: "กลุ่ม" },
  "cb.stats.output": { en: "OUTPUT FILES", th: "ไฟล์ผลลัพธ์" },
  "cb.stats.slots": { en: "SLOTS", th: "ช่อง" },

  // Combo list
  "cb.files.title": { en: "Output Files", th: "ไฟล์ผลลัพธ์" },
  "cb.files.all": { en: "All", th: "ทั้งหมด" },
  "cb.files.none": { en: "None", th: "ไม่เลือก" },
  "cb.preview.click": { en: "Click a combo to preview", th: "คลิกคอมโบเพื่อดูตัวอย่าง" },
  "cb.preview.left": { en: "Left Column", th: "คอลัมน์ซ้าย" },
  "cb.preview.right": { en: "Right Column", th: "คอลัมน์ขวา" },
  "cb.preview.no_right": { en: "No right column", th: "ไม่มีคอลัมน์ขวา" },

  // Settings
  "cb.settings": { en: "Settings", th: "ตั้งค่า" },
  "cb.settings.vgap": { en: "Vertical gap", th: "ระยะห่างแนวตั้ง" },
  "cb.settings.cgap": { en: "Column gap", th: "ระยะห่างคอลัมน์" },

  // Excel preview
  "cb.excel.view": { en: "View order data", th: "ดูข้อมูลคำสั่งซื้อ" },
  "cb.excel.hide": { en: "Hide", th: "ซ่อน" },
  "cb.excel.rows": { en: "rows", th: "แถว" },
  "cb.excel.showing": { en: "Showing", th: "แสดง" },
  "cb.excel.of": { en: "of", th: "จาก" },

  // Export
  "cb.export.btn": { en: "Export", th: "ส่งออก" },
  "cb.export.files": { en: "Files", th: "ไฟล์" },
  "cb.export.file": { en: "File", th: "ไฟล์" },
  "cb.export.exporting": { en: "Combining files", th: "กำลังรวมไฟล์" },
  "cb.export.elapsed": { en: "s elapsed", th: "วินาทีผ่านไป" },
  "cb.export.estimate": { en: "Estimated time:", th: "เวลาโดยประมาณ:" },
  "cb.export.progress": { en: "Server is combining your DST files — please wait, do not close this page", th: "เซิร์ฟเวอร์กำลังรวมไฟล์ DST — กรุณารอ อย่าปิดหน้านี้" },
  "cb.export.downloading": { en: "Downloading zip...", th: "กำลังดาวน์โหลด zip..." },
  "cb.export.done": { en: "Downloaded", th: "ดาวน์โหลดแล้ว" },
  "cb.export.again": { en: "Download again", th: "ดาวน์โหลดอีกครั้ง" },
  "cb.export.previous": { en: "Previously exported", th: "ส่งออกก่อนหน้านี้แล้ว" },
  "cb.export.need_dst": { en: "Upload DST files to enable export", th: "อัปโหลดไฟล์ DST เพื่อเปิดใช้งานการส่งออก" },
  "cb.exported": { en: "Exported", th: "ส่งออกแล้ว" },
  "cb.export.success": { en: "Exported {n} output files", th: "ส่งออก {n} ไฟล์สำเร็จ" },
  "cb.export.completed_in": { en: "Completed in", th: "เสร็จใน" },

  // Error messages
  "err.excel_format": { en: "Please upload an Excel file (.xlsx or .xls)", th: "กรุณาอัปโหลดไฟล์ Excel (.xlsx หรือ .xls)" },
  "err.excel_read": { en: "Failed to read Excel", th: "ไม่สามารถอ่านไฟล์ Excel" },
  "err.excel_drop": { en: "Drop an Excel file (.xlsx or .xls)", th: "ลากไฟล์ Excel (.xlsx หรือ .xls)" },
  "err.no_entries": { en: "No valid entries found with this column mapping. Try adjusting the columns.", th: "ไม่พบข้อมูลที่ถูกต้อง ลองปรับคอลัมน์ใหม่" },
  "err.parse_fail": { en: "Failed to parse Excel", th: "ไม่สามารถอ่านข้อมูล Excel" },
  "err.upload_excel_first": { en: "Upload an Excel order first", th: "อัปโหลดไฟล์ Excel คำสั่งซื้อก่อน" },
  "err.ngs_skipped": { en: "NGS file(s) skipped — convert to DST first", th: "ข้ามไฟล์ NGS — แปลงเป็น DST ก่อน" },
  "err.no_dst": { en: "No DST files found. Upload .dst files or a .zip containing them.", th: "ไม่พบไฟล์ DST อัปโหลดไฟล์ .dst หรือ .zip" },
  "err.no_dst_content": { en: "No DST files found in the uploaded content", th: "ไม่พบไฟล์ DST ในไฟล์ที่อัปโหลด" },
  "err.missing_dst": { en: "Missing {n} DST files", th: "ไม่พบไฟล์ DST {n} ไฟล์" },
  "err.upload_fail": { en: "Upload failed", th: "อัปโหลดไม่สำเร็จ" },
  "err.upload_dst_first": { en: "Upload DST files before exporting", th: "อัปโหลดไฟล์ DST ก่อนส่งออก" },
  "err.export_fail": { en: "Export failed", th: "ส่งออกไม่สำเร็จ" },
  "err.connection": { en: "Connection error", th: "เชื่อมต่อไม่ได้" },
  "err.load_session": { en: "Failed to load session", th: "โหลดเซสชันไม่สำเร็จ" },
  "err.delete_session": { en: "Failed to delete session", th: "ลบเซสชันไม่สำเร็จ" },
  "err.load_sample": { en: "Failed to load sample data", th: "โหลดข้อมูลตัวอย่างไม่สำเร็จ" },

  // Success messages
  "ok.session_deleted": { en: "Session deleted", th: "ลบเซสชันแล้ว" },
  "ok.excel_removed": { en: "Excel data removed", th: "ลบข้อมูล Excel แล้ว" },
  "ok.warnings": { en: "warning(s) during parsing", th: "คำเตือนระหว่างอ่านข้อมูล" },

  // Session picker extras
  "cb.session.no_sessions": { en: "No sessions yet", th: "ยังไม่มีเซสชัน" },
  "cb.session.delete_confirm": { en: "Click again to confirm", th: "คลิกอีกครั้งเพื่อยืนยัน" },
  "cb.session.delete_confirm_label": { en: "✕ Delete?", th: "✕ ลบ?" },
  "cb.session.delete": { en: "Delete session", th: "ลบเซสชัน" },

  // Excel preview table headers
  "cb.table.program": { en: "Program", th: "โปรแกรม" },
  "cb.table.name": { en: "Name", th: "ชื่อ" },
  "cb.table.title": { en: "2nd Line", th: "บรรทัดที่ 2" },
  "cb.table.qty": { en: "Qty", th: "จำนวน" },
  "cb.table.combo": { en: "Combo", th: "คอมโบ" },
  "cb.table.machine": { en: "MA", th: "MA" },
  "cb.table.group": { en: "→ Group", th: "→ กลุ่ม" },

  // Nav
  "nav.signout": { en: "Sign out", th: "ออกจากระบบ" },

  // General
  "names": { en: "names", th: "ชื่อ" },
  "combos": { en: "combos", th: "คอมโบ" },
  "empty": { en: "empty", th: "ว่าง" },
  "Combo": { en: "Combo", th: "คอมโบ" },
  "file": { en: "file", th: "ไฟล์" },
  "files": { en: "files", th: "ไฟล์" },
  "of": { en: "of", th: "จาก" },
  "slots": { en: "slots", th: "ช่อง" },
};

const I18nCtx = createContext<I18nContext>({
  lang: "en",
  toggle: () => {},
  t: (key) => key,
});

export function LanguageProvider({ children }: { children: ReactNode }) {
  const [lang, setLang] = useState<Lang>("en");

  useEffect(() => {
    const saved = localStorage.getItem("lang");
    if (saved === "th" || saved === "en") setLang(saved);
  }, []);

  const toggle = () => {
    const next = lang === "en" ? "th" : "en";
    setLang(next);
    localStorage.setItem("lang", next);
    document.documentElement.lang = next;
  };

  const t = (key: string): string => {
    return translations[key]?.[lang] || key;
  };

  return (
    <I18nCtx.Provider value={{ lang, toggle, t }}>
      {children}
    </I18nCtx.Provider>
  );
}

export function useLanguage() {
  return useContext(I18nCtx);
}

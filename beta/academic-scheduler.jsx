import React, { useState, useMemo, useCallback, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

// Build stamp — injected at build time via esbuild --define:__BUILD__; falls back to "dev" locally.
const BUILD_ID = (typeof __BUILD__ !== "undefined") ? __BUILD__ : "dev";

// ---- Supabase (cloud login + saved schedules) ----
const SUPABASE_URL = "https://afnnrpavqsdamzzkxxcn.supabase.co";
const SUPABASE_KEY = "sb_publishable_aWXZo9kOzMRA5P4J85prBQ_da0jNcVX";
const supabase = (SUPABASE_URL.startsWith("https://") && (SUPABASE_KEY.startsWith("sb_") || SUPABASE_KEY.length > 30))
  ? createClient(SUPABASE_URL, SUPABASE_KEY) : null; // null → app runs fully local, cloud UI hidden
import {
  Calendar, Lock, Unlock, Play, RotateCw, Download, Printer, AlertTriangle,
  CheckCircle2, X, Users, User, DoorOpen, LayoutGrid, Save, Info, ChevronRight,
  Sparkles, Layers, History, Table2, Plus, Trash2, GraduationCap, Gauge, Check, RotateCcw, Upload
} from "lucide-react";

/* ============================================================================
   Academic Schedule Optimizer — interactive mockup
   Placeholder data (Class A · Teacher A). Occupancy is tracked per
   (semester-half × week-parity); biweekly slots render as two triangle blocks
   (top-left = odd weeks, bottom-right = even weeks). Generation is randomized
   and produces several candidate schedules to compare; locked sessions are
   preserved across every candidate. In production the heuristic becomes
   OR-Tools CP-SAT behind this same UI and model.
   ============================================================================ */

const STR = {
  title:{en:"Schedule Optimizer",mn:"Хуваарь зохицуулагч"},
  subtitle:{en:"Accounting · Fall 2026–2027",mn:"Нягтлан бодох бүртгэл · Намрын улирал 2026–2027"},
  buildNote:{en:"Build timestamp — after deploying, hard-refresh and check this changed",mn:"Хувилбарын огноо — байршуулсны дараа шинэчилж, энэ өөрчлөгдсөнийг шалгана уу"},
  quality:{en:"Quality",mn:"Чанар"},
  generate:{en:"Generate",mn:"Хуваарь гаргах"},
  regenerate:{en:"Regenerate",mn:"Дахин гаргах"},
  saveVersion:{en:"Save",mn:"Хадгалах"},
  save:{en:"Save",mn:"Хадгалах"},
  unsavedChanges:{en:"Unsaved changes — click to save a version",mn:"Хадгалаагүй өөрчлөлт — хувилбар хадгалахын тулд дарна уу"},
  saveVersionTitle:{en:"Save a version",mn:"Хувилбар хадгалах"},
  versionName:{en:"Name this version",mn:"Хувилбарын нэр"},
  versionNamePh:{en:"e.g. Fall draft, ready for review",mn:"ж: Намрын ноорог, хянахад бэлэн"},
  versionStatus:{en:"Status",mn:"Төлөв"},
  stInProgress:{en:"In progress",mn:"Хийгдэж байгаа"},
  stFinished:{en:"Finished",mn:"Дууссан"},
  saveVersionBtn:{en:"Save version",mn:"Хувилбар хадгалах"},
  saveOwnerNote:{en:"Saved by {email}",mn:"Хадгалсан: {email}"},
  restoreTitle:{en:"Unsaved changes found",mn:"Хадгалаагүй өөрчлөлт олдлоо"},
  restoreFrom:{en:"From",mn:"Огноо:"},
  restoreBtn:{en:"Restore",mn:"Сэргээх"},
  discardBtn:{en:"Discard",mn:"Устгах"},
  restored:{en:"Unsaved changes restored",mn:"Хадгалаагүй өөрчлөлт сэргээгдлээ"},
  timetable:{en:"Timetable",mn:"Хуваарь"},
  courseData:{en:"Course data",mn:"Хичээлийн мэдээлэл"},
  master:{en:"Master",mn:"Ерөнхий"},
  byCohort:{en:"By cohort",mn:"Бүлгээр"},
  byInstructor:{en:"By instructor",mn:"Багшаар"},
  byRoom:{en:"By room",mn:"Өрөөгөөр"},
  importSummary:{en:"Import summary",mn:"Оруулсан мэдээлэл"},
  firstHalfShort:{en:"1st half",mn:"1-р хагас"},
  secondHalfShort:{en:"2nd half",mn:"2-р хагас"},
  weeks1:{en:"Weeks 1–8",mn:"1–8 долоо хоног"},
  weeks2:{en:"Weeks 9–16",mn:"9–16 долоо хоног"},
  day:{en:"Day",mn:"Өдөр"},
  period:{en:"Period",mn:"Цаг"},
  scheduleQuality:{en:"Schedule quality",mn:"Хуваарийн чанар"},
  hardConstraints:{en:"Hard constraints",mn:"Хатуу нөхцөл"},
  studentExp:{en:"Student experience",mn:"Оюутны туршлага"},
  instructorPref:{en:"Instructor preferences",mn:"Багшийн сонголт"},
  dailyBalance:{en:"Daily balance",mn:"Өдрийн тэнцвэр"},
  roomEff:{en:"Room efficiency",mn:"Өрөөний ашиглалт"},
  needsAttention:{en:"Needs attention",mn:"Анхаарал шаардлагатай"},
  classDetail:{en:"Class detail",mn:"Хичээлийн дэлгэрэнгүй"},
  versions:{en:"Saved versions",mn:"Хадгалсан хувилбар"},
  candidates:{en:"Schedule options",mn:"Хуваарийн сонголт"},
  option:{en:"Option",mn:"Сонголт"},
  pickOption:{en:"Each Generate produces several valid options — click one to compare and choose.",mn:"Хуваарь гаргах бүрд хэд хэдэн сонголт үүснэ — харьцуулж сонгоно уу."},
  parentClass:{en:"Class",mn:"Хичээл"},
  component:{en:"Component",mn:"Төрөл"},
  cohorts:{en:"Cohort(s)",mn:"Бүлэг"},
  colorBy:{en:"Colour",mn:"Өнгө"},
  byCohortC:{en:"Year",mn:"Курс"},
  byTypeC:{en:"Type",mn:"Төрөл"},
  colorByHint:{en:"Colour cards by cohort or by class type (lecture / seminar / lab / bonus)",mn:"Картыг бүлгээр эсвэл хичээлийн төрлөөр (лекц / семинар / лаб / нэмэлт) өнгөөр ялгах"},
  alreadyBooked:{en:"already has a class at that time",mn:"тэр цагт аль хэдийн хичээлтэй"},
  noRoomFree:{en:"no free room fits here",mn:"тохирох чөлөөт өрөө алга"},
  tueBlocked:{en:"No classes after 13:45 on Tuesday",mn:"Мягмар гарагт 13:45-аас хойш хичээл байхгүй"},
  isOffOn:{en:"is off on",mn:"амралттай —"},
  lockHalf:{en:"Lock",mn:"Түгжих"},
  unlockHalf:{en:"Unlock",mn:"Тайлах"},
  lockedHalf:{en:"Half locked — Regenerate now works on the other half",mn:"Хагас түгжигдлээ — Дахин үүсгэх нөгөө хагаст ажиллана"},
  unlockedHalf:{en:"Half unlocked",mn:"Хагас тайлагдлаа"},
  lockHalfHint:{en:"Lock every class in this half so it stays fixed while you build the other half",mn:"Энэ хагасын бүх хичээлийг түгжиж, нөгөө хагасыг зохиох үед хөдөлгөөнгүй байлгана"},
  regenHalfHint:{en:"Re-schedule only this half's unlocked classes (the other half stays fixed)",mn:"Зөвхөн энэ хагасын түгжээгүй хичээлийг дахин байрлуулна (нөгөө хагас хэвээр)"},
  roomPriorityTitle:{en:"Lecture-hall priority",mn:"Лекцийн танхимын дараалал"},
  roomPriorityNote:{en:"When a lecture needs a hall, the optimizer tries them in this order (1 = first choice). E.g. prefer 104 over 306.",mn:"Лекцэд танхим хэрэгтэй үед оновчлогч энэ дарааллаар оролдоно (1 = эхний сонголт). Ж: 306-аас 104-ийг илүүд үзэх."},
  instructor:{en:"Instructor",mn:"Багш"},
  room:{en:"Room",mn:"Өрөө"},
  dayPeriod:{en:"Day / period",mn:"Өдөр / цаг"},
  students:{en:"Students",mn:"Оюутан"},
  lockClass:{en:"Lock class",mn:"Түгжих"},
  unlockClass:{en:"Unlock class",mn:"Тайлах"},
  restore:{en:"Restore",mn:"Сэргээх"},
  online:{en:"Online",mn:"Онлайн"},
  frequency:{en:"Frequency",mn:"Давтамж"},
  length:{en:"Length",mn:"Үргэлжлэх"},
  weekly:{en:"Weekly",mn:"7 хоног бүр"},
  biweekly:{en:"Every 2 weeks",mn:"2 долоо хоногт"},
  full:{en:"Full semester",mn:"Бүтэн улирал"},
  firstHalf:{en:"First half",mn:"Эхний хагас"},
  secondHalf:{en:"Second half",mn:"Сүүлийн хагас"},
  oddWk:{en:"Odd wks",mn:"Сондгой"},
  evenWk:{en:"Even wks",mn:"Тэгш"},
  noVersions:{en:"No saved versions yet. Press Save to store the current schedule as a version you can return to.",mn:"Хадгалсан хувилбар алга. Одоогийн хуваарийг хувилбар болгон хадгалахын тулд Хадгалах дар."},
  renameVersion:{en:"Rename this version:",mn:"Хувилбарын нэрийг өөрчлөх:"},
  delete:{en:"Delete",mn:"Устгах"},
  adminBtn:{en:"Admin",mn:"Админ"},
  adminHint:{en:"Load any user's saved schedule to review or plan together",mn:"Аль ч хэрэглэгчийн хуваарийг ачаалж хамтран төлөвлөх"},
  adminTitle:{en:"Admin — all users' schedules",mn:"Админ — бүх хэрэглэгчийн хуваарь"},
  adminDesc:{en:"Load any user's saved version into your workspace to troubleshoot or plan with them. Your edits save under your own account, never overwriting theirs.",mn:"Аль ч хэрэглэгчийн хадгалсан хувилбарыг ачаалж, тэдэнтэй хамтран засварлана. Таны өөрчлөлт таны бүртгэлд хадгалагдана, тэднийхийг дарж бичихгүй."},
  adminSearch:{en:"Search by email or name…",mn:"Имэйл, нэрээр хайх…"},
  adminEmpty:{en:"No schedules found.",mn:"Хуваарь олдсонгүй."},
  unknownUser:{en:"(unknown user)",mn:"(тодорхойгүй)"},
  refresh:{en:"Refresh",mn:"Шинэчлэх"},
  adminTabSchedules:{en:"Schedules",mn:"Хуваарь"},
  adminTabUsers:{en:"Users",mn:"Хэрэглэгчид"},
  adminUsersDesc:{en:"Everyone who has signed up. Make someone an admin so they can review all schedules and manage users — no need to copy IDs.",mn:"Бүртгүүлсэн бүх хэрэглэгч. Хэн нэгнийг админ болгосноор бүх хуваарийг харах, хэрэглэгч удирдах боломжтой — ID хуулах шаардлагагүй."},
  adminUserSearch:{en:"Search users by email…",mn:"Имэйлээр хайх…"},
  adminBadge:{en:"Admin",mn:"Админ"},
  makeAdmin:{en:"Make admin",mn:"Админ болгох"},
  removeAdmin:{en:"Remove admin",mn:"Админ болиулах"},
  madeAdmin:{en:"Admin access granted",mn:"Админ эрх олгосон"},
  removedAdmin:{en:"Admin access removed",mn:"Админ эрх хассан"},
  joined:{en:"Joined",mn:"Элссэн"},
  you:{en:"you",mn:"та"},
  capacityTitle:{en:"Group over capacity",mn:"Бүлгийн ачаалал хэтэрсэн"},
  repeatsTitle:{en:"Repeated session same day",mn:"Нэг өдөрт давхардсан хичээл"},
  conflictsTitle:{en:"Scheduling conflicts",mn:"Хуваарийн зөрчил"},
  conflictsDesc:{en:"Two classes share the same slot in the same week — a real clash. New schedules never create these; any shown came from an older saved version or a manual move. Regenerate, or drag one class to a free slot.",mn:"Хоёр хичээл нэг долоо хоногт нэг цагт давхцаж байна — жинхэнэ зөрчил. Шинэ хуваарь эдгээрийг үүсгэхгүй; эндхийн зүйлс хуучин хадгалалт эсвэл гараар зөөсөнөөс болсон. Дахин үүсгэх, эсвэл нэг хичээлийг чөлөөт цаг руу чирнэ үү."},
  cf_group:{en:"same group",mn:"нэг бүлэг"},
  cf_instructor:{en:"same instructor",mn:"нэг багш"},
  cf_room:{en:"same room",mn:"нэг өрөө"},
  repeatsDesc:{en:"Lecture + seminar or lecture + lab on the same day is fine — this lists only two of the SAME type (e.g. two seminars) for one group in one day, which is blocked during generation. Any shown here are locked/manually-placed: free a slot, split across instructors, or drag one to another day.",mn:"Лекц + семинар эсвэл лекц + лаб нэг өдөрт байж болно — энд зөвхөн ижил төрлийн хоёр хичээл (ж: хоёр семинар) нэг бүлэгт нэг өдөрт орсныг харуулна, үүнийг үүсгэх үед хориглоно. Эндхийн зүйлс түгжсэн/гараар байрлуулсан: цаг чөлөөлөх, багш хуваах, эсвэл өөр өдөрт чирнэ үү."},
  sameTypeDay:{en:"Two of the same session type for this group on one day",mn:"Энэ бүлэгт нэг өдөрт ижил төрлийн хоёр хичээл"},
  capacityDesc:{en:"These groups are at or over the weekly limit ({n} class slots per half). When a group is this full, some classes stay unplaced or must use discouraged slots (P4 / Tuesday afternoon). Reduce the group's load or split it across instructors.",mn:"Эдгээр бүлэг долоо хоногийн хязгаарт хүрсэн эсвэл хэтэрсэн (хагас бүрт {n} цаг). Ийм дүүрэн үед зарим хичээл байрлахгүй эсвэл тохиромжгүй цаг (P4 / Мягмар үдээс хойш) ашиглана. Ачааллыг бууруулах эсвэл багш хуваана уу."},
  signInSaveVersions:{en:"Sign in to save versions you can return to (across sessions and devices).",mn:"Хувилбар хадгалахын тулд нэвтэрнэ үү (сесс, төхөөрөмж хооронд)."},
  unplacedHint:{en:"No feasible slot with a free room + instructor. Unlock nearby classes or add a room/period.",mn:"Багш, өрөө сул зэрэгцэх цаг олдсонгүй. Ойролцоох хичээлийг тайлах эсвэл өрөө/цаг нэмнэ үү."},
  dragHint:{en:"Drag a class to an empty slot — hard conflicts are blocked live. The late 13:50 period warns but is allowed; Tuesday afternoon is blocked. Split cells alternate by week: top-left = odd, bottom-right = even.",mn:"Хичээлийг сул нүд рүү чирнэ үү — хатуу зөрчлийг шууд хориглоно. 13:50 оройн цаг сануулна ч болно; Мягмар үдээс хойш хаалттай. Хуваасан нүд долоо хоногоор ээлжилнэ: зүүн дээд = сондгой, баруун доод = тэгш."},
  importWhat:{en:"This is what the app parsed from your Excel import. Resolve any flags below before trusting the generated schedule.",mn:"Энэ бол Excel-ээс уншсан мэдээлэл. Хуваарийг найдахаас өмнө доорх анхааруулгыг шийднэ үү."},
  cohortsN:{en:"cohorts",mn:"бүлэг"},
  coursesN:{en:"courses",mn:"хичээл"},
  periodsWk:{en:"periods/wk",mn:"цаг/долоо хоног"},
  instructorsN:{en:"instructors",mn:"багш"},
  roomsN:{en:"rooms",mn:"өрөө"},
  sessionsN:{en:"session requirements",mn:"хичээлийн цаг"},
  dayOffNote:{en:"1 instructor has a fixed day off",mn:"1 багш тогтмол амралттай"},
  attentionN:{en:"session(s) could not be placed",mn:"хичээл байршуулж чадсангүй"},
  year:{en:"Year",mn:"Курс"},
  gaps:{en:"Mid-day gaps",mn:"Дунд завсар"},
  p4used:{en:"Period-4 classes",mn:"4-р цагийн хичээл"},
  tueLate:{en:"Tuesday after 13:45",mn:"Мягмар 13:45-аас хойш"},
  moved:{en:"Moved — schedule still valid",mn:"Зөөгдлөө — хуваарь хүчинтэй"},
  lockedMove:{en:"Locked class — unlock to move it",mn:"Түгжээтэй хичээл — эхлээд тайлна уу"},
  wrongColumn:{en:"Can't move a class into another group's column — drop it in its own column (a different period or day)",mn:"Хичээлийг өөр бүлгийн баганад зөөх боломжгүй — өөрийнх нь баганад (өөр цаг эсвэл өдөр) буулгана уу"},
  signIn:{en:"Sign in",mn:"Нэвтрэх"},
  signUp:{en:"Sign up",mn:"Бүртгүүлэх"},
  signOut:{en:"Sign out",mn:"Гарах"},
  mySchedules:{en:"My schedules",mn:"Миний хуваарь"},
  cloudTitle:{en:"Cloud schedules",mn:"Үүлэн хуваарь"},
  emailLbl:{en:"Email",mn:"Имэйл"},
  passwordLbl:{en:"Password",mn:"Нууц үг"},
  saveToCloud:{en:"Save to cloud",mn:"Үүлэнд хадгалах"},
  scheduleName:{en:"Schedule name",mn:"Хуваарийн нэр"},
  statusDraft:{en:"Draft",mn:"Ноорог"},
  statusCheckpoint:{en:"In progress",mn:"Хийгдэж байгаа"},
  statusFinal:{en:"Finished",mn:"Дууссан"},
  statusAutosave:{en:"Auto-saved",mn:"Авто-хадгалсан"},
  loadBtn:{en:"Load",mn:"Ачаалах"},
  noSaved:{en:"No saved schedules yet.",mn:"Хадгалсан хуваарь алга."},
  cloudSaved:{en:"Saved to cloud",mn:"Үүлэнд хадгалагдлаа"},
  cloudLoaded:{en:"Loaded from cloud",mn:"Үүлнээс ачааллаа"},
  cloudErr:{en:"Cloud error",mn:"Үүлэн алдаа"},
  signInToSave:{en:"Sign in to save",mn:"Хадгалахын тулд нэвтэрнэ үү"},
  authHint:{en:"New here? Enter an email + password and press Sign up. Otherwise Sign in.",mn:"Шинэ хэрэглэгч үү? Имэйл, нууц үг оруулаад Бүртгүүлэх дар. Үгүй бол Нэвтрэх."},
  savedAs:{en:"Saved as Version",mn:"Хадгалсан: Хувилбар"},
  restoredV:{en:"Restored Version",mn:"Сэргээв: Хувилбар"},
  solverNote:{en:"CP-SAT in production · heuristic solver in this mockup",mn:"Бодит хувилбарт CP-SAT · энэ загварт эвристик"},
  gen1:{en:"Validating data…",mn:"Мэдээлэл шалгаж байна…"},
  gen2:{en:"Building session requirements…",mn:"Хичээлийн цаг бүрдүүлж байна…"},
  gen3:{en:"Exploring feasible schedules…",mn:"Боломжит хуваарь хайж байна…"},
  gen4:{en:"Optimizing preferences…",mn:"Сайжруулж байна…"},
  gen5:{en:"Ranking options…",mn:"Сонголт эрэмбэлж байна…"},
  editHint:{en:"This is the normalized data from your Excel import. Edit it here, then apply to regenerate.",mn:"Энэ бол Excel-ээс оруулсан мэдээлэл. Эндээс засаад дахин хуваарь гаргана уу."},
  colName:{en:"Course name",mn:"Хичээлийн нэр"},
  colTeacher:{en:"Teacher",mn:"Багш"},
  colCohorts:{en:"Cohorts",mn:"Бүлэг"},
  colLength:{en:"Length",mn:"Урт"},
  colCredits:{en:"Credit hours by room type",mn:"Кредит (өрөөний төрлөөр)"},
  colGenerates:{en:"Generates",mn:"Үүсгэх"},
  cLecture:{en:"Lecture",mn:"Лекц"},
  cSeminar:{en:"Seminar",mn:"Семинар"},
  cBonus:{en:"Bonus",mn:"Нэмэлт"},
  cLab:{en:"Lab",mn:"Лаб"},
  cPractical:{en:"Practical",mn:"Дадлага"},
  addCourse:{en:"Add class",mn:"Хичээл нэмэх"},
  addTeacher:{en:"Add teacher",mn:"Багш нэмэх"},
  deleteTeacher:{en:"Delete teacher",mn:"Багш устгах"},
  cantDeleteTeacher:{en:"Can't delete — this instructor still has assigned courses. Reassign or remove them first.",mn:"Устгах боломжгүй — энэ багшид оногдсон хичээл байна. Эхлээд өөр багшид шилжүүлнэ үү."},
  importXlsx:{en:"Import Excel",mn:"Excel оруулах"},
  exportXlsx:{en:"Download data",mn:"Өгөгдөл татах"},
  importOk:{en:"Imported",mn:"Оруулсан"},
  importErr:{en:"Couldn't read that file. Make sure it's the course-data template (Courses + Instructors sheets).",mn:"Файлыг уншиж чадсангүй. Хичээлийн өгөгдлийн загвар (Courses + Instructors) эсэхийг шалгана уу."},
  teachersTitle:{en:"Teachers & availability",mn:"Багш ба бэлэн байдал"},
  dayOff:{en:"Day off",mn:"Амралт"},
  none:{en:"None",mn:"Байхгүй"},
  applyRegen:{en:"Apply & regenerate",mn:"Хадгалаад дахин гаргах"},
  creditRuleNote:{en:"Rule: 1 lecture credit = 8 periods · 1 seminar/lab/practical credit = 16 periods. A full-semester 1-credit lecture becomes biweekly.",mn:"Дүрэм: 1 лекц кредит = 8 цаг · 1 семинар/лаб/дадлага кредит = 16 цаг. Бүтэн улирлын 1 кредит лекц нь 2 долоо хоногт болно."},
  sharedLecture:{en:"Shared lecture",mn:"Хамтарсан лекц"},
  sharedSeminar:{en:"Shared seminar",mn:"Хамтарсан семинар"},
  onlineFlag:{en:"Online",mn:"Онлайн"},
  parallelFlag:{en:"Parallel (2 rooms/teachers)",mn:"Зэрэгцээ (2 өрөө/багш)"},
  parallelHint:{en:"Split the seminar into 2 level sections that run at the same time in 2 rooms with 2 teachers (e.g. English by level). Rare.",mn:"Семинарыг түвшингээр 2 хэсэг болгож, нэг цагт 2 өрөөнд 2 багшаар явуулна (ж: түвшингээр ангилсан англи хэл). Ховор."},
  parallelNote:{en:"Parallel — 2 level sections, 2 rooms, 2 teachers, same time",mn:"Зэрэгцээ — 2 түвшин, 2 өрөө, 2 багш, нэг цагт"},
  teacher2:{en:"2nd teacher…",mn:"2-р багш…"},
  section:{en:"Section",mn:"Хэсэг"},
  onlineHint:{en:"Held online — no room is assigned for any part of this course",mn:"Онлайнаар явагдана — энэ хичээлд өрөө хуваарилахгүй"},
  semesterStart:{en:"Semester start date",mn:"Улирлын эхлэх огноо"},
  weekStartsOn:{en:"Week starts on",mn:"Долоо хоног эхлэх"},
  startDateNote:{en:"The teaching week is built from this day. Blank = Monday. Change it, then Generate to re-solve.",mn:"Хичээлийн долоо хоног энэ өдрөөс эхэлнэ. Хоосон = Даваа. Өөрчилсний дараа дахин үүсгэнэ үү."},
  softRequests:{en:"Time requests",mn:"Цагийн хүсэлт"},
  anyDay:{en:"Any day",mn:"Аль ч өдөр"},
  addRequest:{en:"Add",mn:"Нэмэх"},
  requestNote:{en:"Instructors usually dislike a period rather than a day — e.g. no 1st period (school pickup) or no 2nd period. Pick “Any day” to block a period all week; the optimizer avoids it when it can.",mn:"Багш нар ихэвчлэн өдрөөс илүү цагаас татгалздаг — ж: 1-р цаггүй (хүүхэд авах) эсвэл 2-р цаггүй. Тухайн цагийг долоо хоногийн турш хаахын тулд “Аль ч өдөр”-ийг сонго; оновчлогч боломжтой бол зайлсхийнэ."},
  rulesTitle:{en:"Scheduling rules",mn:"Хуваарийн дүрэм"},
  rulesBtn:{en:"Rules",mn:"Дүрэм"},
  hardRules:{en:"Hard rules — always enforced",mn:"Хатуу дүрэм — үргэлж баримтална"},
  softRules:{en:"Soft rules — toggle on/off, then regenerate",mn:"Зөөлөн дүрэм — асаах/унтраах, дараа нь дахин үүсгэх"},
  rulesHint:{en:"Turn a preference off to let the optimizer relax it. Changes apply on the next Generate.",mn:"Аль нэг зөөлөн дүрмийг унтраавал оновчлогч түүнийг сулруулна. Дараагийн үүсгэлтэд хэрэгжинэ."},
  resetSample:{en:"Reset to sample",mn:"Жишээ рүү буцаах"},
  clearAll:{en:"Clear all",mn:"Бүгдийг арилгах"},
  confirmClear:{en:"Clear all classes and start from an empty template? (Use “Reset to sample” to bring the example back.)",mn:"Бүх хичээлийг устгаж хоосон эхлэх үү? (“Жишээ рүү буцаах” товчоор жишээг эргүүлж авна.)"},
  cohortNote:{en:"Click cohort tags to set who a class is for — any combination is allowed. With 2+ cohorts, “Shared lecture” holds one combined lecture in a hall (seminars stay separate per cohort).",mn:"Хичээл хэнд зориулсныг тагаар сонгоно — ямар ч хослол болно. 2+ бүлэгтэй үед “Хамтарсан лекц” нь танхимд нэг лекц болгоно (семинар бүлэг бүрт тусдаа)."},
  teacherRating:{en:"Teacher schedule rating",mn:"Багшийн хуваарийн үнэлгээ"},
  workload:{en:"Workload",mn:"Ачаалал"},
  breaksLbl:{en:"Breaks / gaps",mn:"Завсарлага"},
  consecutive:{en:"Consecutive classes",mn:"Дараалсан хичээл"},
  allTeachers:{en:"All teachers (lowest first)",mn:"Бүх багш (доогуураас)"},
  classesN:{en:"classes",mn:"хичээл"},
  daysN:{en:"days",mn:"өдөр"},
};
const tr = (lang, k) => (STR[k] ? (STR[k][lang] ?? STR[k].en) : k);

// ---------- Time ----------
// Canonical weekdays (Mon–Fri). The teaching week is rotated to begin on the
// semester's start weekday, derived from the start date entered in Course Data.
const WEEKDAYS = [
  { id:"mon", en:"Monday", mn:"Даваа" }, { id:"tue", en:"Tuesday", mn:"Мягмар" },
  { id:"wed", en:"Wednesday", mn:"Лхагва" }, { id:"thu", en:"Thursday", mn:"Пүрэв" },
  { id:"fri", en:"Friday", mn:"Баасан" },
];
const DEFAULT_TERM_START = "2026-09-01"; // Fall — a Tuesday; blank falls back to Monday
function weekdayIndex(dateStr) { // 0=Mon … 4=Fri, weekend→Mon (0)
  if (!dateStr) return 0;
  const d = new Date(dateStr + "T00:00:00");
  if (isNaN(d.getTime())) return 0;
  const idx = (d.getDay() + 6) % 7; // JS 0=Sun → 0=Mon
  return idx > 4 ? 0 : idx;
}
function weekOrderFromStart(dateStr) {
  const i = weekdayIndex(dateStr);
  return WEEKDAYS.slice(i).concat(WEEKDAYS.slice(0, i));
}
// `DAYS` holds the active teaching-week order; App reassigns it from the start date.
let DAYS = weekOrderFromStart(DEFAULT_TERM_START);
const dayLabel = (lang, d) => (lang === "mn" ? d.mn : d.en);
const PERIODS = [
  { id:1, label:"08:40–10:10" }, { id:2, label:"10:15–11:45" },
  { id:3, label:"12:15–13:45" }, { id:4, label:"13:50–15:20" },
];
const TUE_CUTOFF_PERIOD = 4;
const slotIndex = (dayId, period) => DAYS.findIndex((d) => d.id === dayId) * 4 + (period - 1);

// ---------- Naming (letters) ----------
const letter = (n) => { let s=""; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26);} return s; };
const defClassName = (idx) => `Class ${letter(idx)}`;
const defTeacherName = (n) => `Teacher ${letter(n)}`;

// ---------- Components ----------
const COMP_FULL = {
  L:{en:"Lecture",mn:"Лекц"}, S:{en:"Seminar",mn:"Семинар"}, B:{en:"Bonus seminar",mn:"Нэмэлт семинар"},
  Lab:{en:"Laboratory",mn:"Лаборатори"}, P:{en:"Practical",mn:"Дадлага"},
};
const COMP_TAG = { L:{en:"L",mn:"Л"}, S:{en:"S",mn:"С"}, B:{en:"B",mn:"Б"}, Lab:{en:"Lab",mn:"Лаб"}, P:{en:"Pr",mn:"Дад"} };
const COMP_PHASE = { L:0, S:1, Lab:1, P:1, B:2 };
const compLabel = (lang,t) => COMP_FULL[t][lang];
const compTag = (lang,t) => COMP_TAG[t][lang];

// ---------- Cohorts / rooms ----------
// Colour by year level (well-separated hues; A/B share their year's colour)
const YEAR_COLORS = { 1:"#7c3aed", 2:"#16a34a", 3:"#ea580c", 4:"#dc2626" }; // purple, green, orange, red
const cohortColor = (id) => YEAR_COLORS[cohort(id)?.year] || "#64748b";
const COHORT_COLORS = new Proxy({}, { get:(_,k)=> cohortColor(String(k)) }); // back-compat: any COHORT_COLORS[id] → year colour
// Class-type colours reuse the same 4-hue palette (lecture→purple, seminar→green, bonus→orange, lab→red)
const COMP_COLORS = { L:"#7c3aed", S:"#16a34a", B:"#ea580c", Lab:"#dc2626", P:"#0891b2" };
const cardColor = (s, colorBy) => colorBy === "type" ? (COMP_COLORS[s.type] || "#64748b") : (YEAR_COLORS[cohort(s.cohorts[0])?.year] || "#64748b");
const COHORTS = [
  { id:"1A", year:1, students:24, home:"304" }, { id:"1B", year:1, students:24, home:"309" },
  { id:"2A", year:2, students:20, home:"308" }, { id:"2B", year:2, students:20, home:"305" },
  { id:"3A", year:3, students:20, home:"303" }, { id:"4A", year:4, students:18, home:"307" },
];
const cohort = (id) => COHORTS.find((c) => c.id === id);
const ROOMS = [
  { id:"304", type:"seminar", cap:24 }, { id:"309", type:"seminar", cap:24 },
  { id:"308", type:"seminar", cap:24 }, { id:"305", type:"seminar", cap:24 },
  { id:"303", type:"seminar", cap:24 }, { id:"307", type:"seminar", cap:24 },
  { id:"306", type:"lecture", cap:50 }, { id:"104", type:"lecture", cap:50 },
  { id:"201", type:"lab", cap:24 },
];
const LECTURE_HALLS = ["306","104"];
// Lecture-hall preference (lower = tried first). App reassigns from the Room priority editor. 104 has nicer chairs → default first.
let ROOM_PRI = { "104":1, "306":2 };
const hallsByPri = () => [...LECTURE_HALLS].sort((a,b)=>(ROOM_PRI[a]??50)-(ROOM_PRI[b]??50));
const homeRoom = (id) => cohort(id)?.home;

// ---------- Seed data (editable) ----------
const mkTeacher = (i, off=null, pref=null) => ({ id:`t${i}`, name:defTeacherName(i), off, pref });
const SEED_TEACHERS = [
  {id:"t1", name:"Ё.Баяртайван", off:null, pref:null},
  {id:"t2", name:"Ц.Батсүрэн", off:null, pref:null},
  {id:"t3", name:"Б.Сайханцэцэг", off:null, pref:null},
  {id:"t4", name:"С.Мөнхцэцэг", off:null, pref:null},
  {id:"t5", name:"А.Шижирбаатар", off:null, pref:null},
  {id:"t6", name:"Б.Соджавхлан", off:null, pref:null},
  {id:"t7", name:"Шинэ багш 1", off:null, pref:null},
  {id:"t8", name:"Б.Мөнхбаяр", off:null, pref:null},
  {id:"t9", name:"Ц.Энхбилэг", off:null, pref:null},
  {id:"t10", name:"Шинэ багш 2", off:null, pref:null},
  {id:"t11", name:"Ж.Бадамханд", off:null, pref:null},
  {id:"t12", name:"М.Ганхөлөг", off:null, pref:null},
  {id:"t13", name:"Хуваарилаагүй 1", off:null, pref:null},
  {id:"t14", name:"Б.Мэндбаяр", off:null, pref:null},
  {id:"t15", name:"Х.Одмандах", off:null, pref:null},
  {id:"t16", name:"Англи багш 2", off:null, pref:null},
];
const cc = (idx, cohorts, t, length, o) => ({
  idx, name:defClassName(idx), cohorts, t, length,
  lec:0, sem:0, bonus:0, lab:0, prac:0, online:false, combineLec:false, combineSem:false, parallel:false, t2:null, ...o,
});
const SEED_COURSES = [
  cc(1, ["1A","1B"], "t1", "h1", { name:"Математик", lec:1, sem:0.5, combineLec:true }),
  cc(2, ["1A","1B"], "t2", "full", { name:"Мэдээллийн технологийн ашиглалт", lec:1, lab:1, combineLec:true }),
  cc(3, ["1A","1B"], "t3", "full", { name:"Монголын түүх, соёл", lec:2, sem:1, combineLec:true }),
  cc(4, ["1A","1B"], "t4", "full", { name:"Микро экономикс", lec:2, sem:1, online:true, combineLec:true, combineSem:true }),
  cc(5, ["1A","1B"], "t5", "full", { name:"Биеийн тамир, эрүүл мэнд", sem:1, combineLec:true }),
  cc(6, ["1A","1B"], "t6", "full", { name:"Монгол хэлний найруулга зүй", lec:2, sem:1, combineLec:true }),
  cc(7, ["1A","1B"], "t7", "full", { name:"Бизнесийн харилцаа, ёс зүй", lec:2, sem:1, combineLec:true }),
  cc(8, ["1A","1B"], "t8", "full", { name:"НББ-ийн үндэс", lec:2, sem:1, bonus:1, combineLec:true }),
  cc(9, ["2A","2B"], "t9", "full", { name:"Санхүүгийн үндэс", lec:2, sem:1, combineLec:true }),
  cc(10, ["2A","2B"], "t10", "full", { name:"Менежментийн үндэс", lec:2, sem:1, combineLec:true }),
  cc(11, ["2A","2B"], "t11", "full", { name:"Судалгааны арга зүйн үндэс", lec:2, sem:1, combineLec:true }),
  cc(12, ["2A","2B"], "t12", "full", { name:"Бизнесийн эрх зүй", lec:2, sem:1, bonus:1, online:true, combineLec:true, combineSem:true }),
  cc(13, ["2A","2B"], "t11", "h1", { name:"Санхүүгийн тайлагнал ба мэдээллийн систем 1", lec:1, sem:0.5, bonus:0.5, combineLec:true }),
  cc(14, ["2A","2B"], "t13", "h2", { name:"Санхүүгийн тайлагнал ба мэдээллийн систем 2", lec:1, lab:1, combineLec:true }),
  cc(15, ["2A","2B"], "t6", "full", { name:"Бизнесийн англи хэл", sem:3, parallel:true, t2:"t16" }),
  cc(16, ["3A"], "t5", "full", { name:"СДШНББ II", lec:2, sem:1, bonus:1 }),
  cc(17, ["3A"], "t8", "full", { name:"Аудит ба баталгаажуулалт", lec:2, sem:1 }),
  cc(18, ["3A"], "t5", "full", { name:"Татварын бүртгэл", lec:2, sem:1 }),
  cc(19, ["3A"], "t6", "full", { name:"Мэргэжлийн англи хэл II", sem:3 }),
  cc(20, ["3A"], "t14", "full", { name:"Өртгийн бүртгэл", lec:2, sem:1, bonus:1 }),
  cc(21, ["3A"], "t14", "h2", { name:"Дижитал нягтлан бодох бүртгэл", lec:2, lab:1 }),
  cc(22, ["4A"], "t14", "h1", { name:"Өртгийн удирдлага", lec:2, sem:1, bonus:1 }),
  cc(23, ["4A"], "t15", "h1", { name:"Санхүүгийн тайлагналын ОУС", lec:2, sem:1 }),
  cc(24, ["4A"], "t9", "h1", { name:"Санхүүгийн шинжилгээ, шийдвэр гаргалт", lec:2, sem:1 }),
  cc(25, ["4A"], "t8", "h1", { name:"НББ-ийн мэргэжлийн судалгаа", sem:0.25 }),
];
const SEED_PLACED = [
  {id:"P0",courseIdx:22,ins:"t14",type:"L",roomType:"lecture",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"wed",period:1,room:"307",parity:"weekly",locked:false},
  {id:"P1",courseIdx:22,ins:"t14",type:"L",roomType:"lecture",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"tue",period:1,room:"307",parity:"weekly",locked:false},
  {id:"P2",courseIdx:22,ins:"t14",type:"S",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"thu",period:1,room:"307",parity:"weekly",locked:false},
  {id:"P3",courseIdx:22,ins:"t14",type:"S",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"fri",period:1,room:"307",parity:"weekly",locked:false},
  {id:"P4",courseIdx:22,ins:"t14",type:"B",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"mon",period:1,room:"307",parity:"weekly",locked:false},
  {id:"P5",courseIdx:22,ins:"t14",type:"B",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"fri",period:2,room:"307",parity:"weekly",locked:false},
  {id:"P6",courseIdx:8,ins:"t8",type:"L",roomType:"lecture",freq:"weekly",phase:"full",cohorts:["1A","1B"],students:48,day:"fri",period:1,room:"104",parity:"weekly",locked:false},
  {id:"P7",courseIdx:8,ins:"t8",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1A"],students:24,day:"mon",period:1,room:"304",parity:"weekly",locked:false},
  {id:"P8",courseIdx:8,ins:"t8",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1B"],students:24,day:"fri",period:2,room:"309",parity:"weekly",locked:false},
  {id:"P9",courseIdx:8,ins:"t8",type:"B",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1A"],students:24,day:"mon",period:2,room:"304",parity:"weekly",locked:false},
  {id:"P10",courseIdx:8,ins:"t8",type:"B",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1B"],students:24,day:"mon",period:3,room:"309",parity:"weekly",locked:false},
  {id:"P11",courseIdx:13,ins:"t11",type:"L",roomType:"lecture",freq:"weekly",phase:"h1",cohorts:["2A","2B"],students:40,day:"fri",period:1,room:"306",parity:"weekly",locked:false},
  {id:"P12",courseIdx:13,ins:"t11",type:"S",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["2A"],students:20,day:"mon",period:1,room:"308",parity:"weekly",locked:false},
  {id:"P13",courseIdx:13,ins:"t11",type:"S",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["2B"],students:20,day:"fri",period:2,room:"305",parity:"weekly",locked:false},
  {id:"P14",courseIdx:13,ins:"t11",type:"B",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["2A"],students:20,day:"mon",period:2,room:"308",parity:"weekly",locked:false},
  {id:"P15",courseIdx:13,ins:"t11",type:"B",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["2B"],students:20,day:"mon",period:3,room:"305",parity:"weekly",locked:false},
  {id:"P16",courseIdx:14,ins:"t13",type:"L",roomType:"lecture",freq:"weekly",phase:"h2",cohorts:["2A","2B"],students:40,day:"tue",period:1,room:"104",parity:"weekly",locked:false},
  {id:"P17",courseIdx:14,ins:"t13",type:"Lab",roomType:"lab",freq:"weekly",phase:"h2",cohorts:["2A"],students:20,day:"thu",period:1,room:"201",parity:"weekly",locked:false},
  {id:"P18",courseIdx:14,ins:"t13",type:"Lab",roomType:"lab",freq:"weekly",phase:"h2",cohorts:["2B"],students:20,day:"mon",period:1,room:"201",parity:"weekly",locked:false},
  {id:"P19",courseIdx:14,ins:"t13",type:"Lab",roomType:"lab",freq:"weekly",phase:"h2",cohorts:["2A"],students:20,day:"fri",period:1,room:"201",parity:"weekly",locked:false},
  {id:"P20",courseIdx:14,ins:"t13",type:"Lab",roomType:"lab",freq:"weekly",phase:"h2",cohorts:["2B"],students:20,day:"wed",period:1,room:"201",parity:"weekly",locked:false},
  {id:"P21",courseIdx:21,ins:"t14",type:"L",roomType:"lecture",freq:"weekly",phase:"h2",cohorts:["3A"],students:20,day:"tue",period:1,room:"303",parity:"weekly",locked:false},
  {id:"P22",courseIdx:21,ins:"t14",type:"L",roomType:"lecture",freq:"weekly",phase:"h2",cohorts:["3A"],students:20,day:"fri",period:1,room:"303",parity:"weekly",locked:false},
  {id:"P23",courseIdx:21,ins:"t14",type:"Lab",roomType:"lab",freq:"weekly",phase:"h2",cohorts:["3A"],students:20,day:"mon",period:2,room:"201",parity:"weekly",locked:false},
  {id:"P24",courseIdx:21,ins:"t14",type:"Lab",roomType:"lab",freq:"weekly",phase:"h2",cohorts:["3A"],students:20,day:"fri",period:2,room:"201",parity:"weekly",locked:false},
  {id:"P25",courseIdx:23,ins:"t15",type:"L",roomType:"lecture",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"thu",period:2,room:"307",parity:"weekly",locked:false},
  {id:"P26",courseIdx:23,ins:"t15",type:"L",roomType:"lecture",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"tue",period:2,room:"307",parity:"weekly",locked:false},
  {id:"P27",courseIdx:23,ins:"t15",type:"S",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"mon",period:2,room:"307",parity:"weekly",locked:false},
  {id:"P28",courseIdx:23,ins:"t15",type:"S",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"thu",period:3,room:"307",parity:"weekly",locked:false},
  {id:"P29",courseIdx:24,ins:"t9",type:"L",roomType:"lecture",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"wed",period:2,room:"307",parity:"weekly",locked:false},
  {id:"P30",courseIdx:24,ins:"t9",type:"L",roomType:"lecture",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"fri",period:3,room:"307",parity:"weekly",locked:false},
  {id:"P31",courseIdx:24,ins:"t9",type:"S",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"mon",period:3,room:"307",parity:"weekly",locked:false},
  {id:"P32",courseIdx:24,ins:"t9",type:"S",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["4A"],students:18,day:"fri",period:4,room:"307",parity:"weekly",locked:false},
  {id:"P33",courseIdx:1,ins:"t1",type:"L",roomType:"lecture",freq:"weekly",phase:"h1",cohorts:["1A","1B"],students:48,day:"tue",period:1,room:"104",parity:"weekly",locked:false},
  {id:"P34",courseIdx:1,ins:"t1",type:"S",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["1A"],students:24,day:"wed",period:1,room:"304",parity:"weekly",locked:false},
  {id:"P35",courseIdx:1,ins:"t1",type:"S",roomType:"seminar",freq:"weekly",phase:"h1",cohorts:["1B"],students:24,day:"thu",period:1,room:"309",parity:"weekly",locked:false},
  {id:"P36",courseIdx:2,ins:"t2",type:"L",roomType:"lecture",freq:"biweekly",phase:"full",cohorts:["1A","1B"],students:48,day:"tue",period:2,room:"104",parity:"even",locked:false},
  {id:"P37",courseIdx:2,ins:"t2",type:"Lab",roomType:"lab",freq:"weekly",phase:"full",cohorts:["1A"],students:24,day:"wed",period:2,room:"201",parity:"weekly",locked:false},
  {id:"P38",courseIdx:2,ins:"t2",type:"Lab",roomType:"lab",freq:"weekly",phase:"full",cohorts:["1B"],students:24,day:"thu",period:2,room:"201",parity:"weekly",locked:false},
  {id:"P39",courseIdx:3,ins:"t3",type:"L",roomType:"lecture",freq:"weekly",phase:"full",cohorts:["1A","1B"],students:48,day:"tue",period:3,room:"104",parity:"weekly",locked:false},
  {id:"P40",courseIdx:3,ins:"t3",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1A"],students:24,day:"thu",period:1,room:"304",parity:"weekly",locked:false},
  {id:"P41",courseIdx:3,ins:"t3",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1B"],students:24,day:"mon",period:2,room:"309",parity:"weekly",locked:false},
  {id:"P42",courseIdx:6,ins:"t6",type:"L",roomType:"lecture",freq:"weekly",phase:"full",cohorts:["1A","1B"],students:48,day:"wed",period:3,room:"104",parity:"weekly",locked:false},
  {id:"P43",courseIdx:6,ins:"t6",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1A"],students:24,day:"fri",period:2,room:"304",parity:"weekly",locked:false},
  {id:"P44",courseIdx:6,ins:"t6",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1B"],students:24,day:"thu",period:3,room:"309",parity:"weekly",locked:false},
  {id:"P45",courseIdx:7,ins:"t7",type:"L",roomType:"lecture",freq:"weekly",phase:"full",cohorts:["1A","1B"],students:48,day:"fri",period:3,room:"104",parity:"weekly",locked:false},
  {id:"P46",courseIdx:7,ins:"t7",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1A"],students:24,day:"mon",period:3,room:"304",parity:"weekly",locked:false},
  {id:"P47",courseIdx:7,ins:"t7",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1B"],students:24,day:"mon",period:1,room:"309",parity:"weekly",locked:false},
  {id:"P48",courseIdx:9,ins:"t9",type:"L",roomType:"lecture",freq:"weekly",phase:"full",cohorts:["2A","2B"],students:40,day:"tue",period:2,room:"306",parity:"weekly",locked:false},
  {id:"P49",courseIdx:9,ins:"t9",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["2A"],students:20,day:"wed",period:1,room:"308",parity:"weekly",locked:false},
  {id:"P50",courseIdx:9,ins:"t9",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["2B"],students:20,day:"mon",period:2,room:"305",parity:"weekly",locked:false},
  {id:"P51",courseIdx:10,ins:"t10",type:"L",roomType:"lecture",freq:"weekly",phase:"full",cohorts:["2A","2B"],students:40,day:"wed",period:2,room:"104",parity:"weekly",locked:false},
  {id:"P52",courseIdx:10,ins:"t10",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["2A"],students:20,day:"fri",period:2,room:"308",parity:"weekly",locked:false},
  {id:"P53",courseIdx:10,ins:"t10",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["2B"],students:20,day:"thu",period:2,room:"305",parity:"weekly",locked:false},
  {id:"P54",courseIdx:11,ins:"t11",type:"L",roomType:"lecture",freq:"weekly",phase:"full",cohorts:["2A","2B"],students:40,day:"tue",period:3,room:"306",parity:"weekly",locked:false},
  {id:"P55",courseIdx:11,ins:"t11",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["2A"],students:20,day:"thu",period:2,room:"308",parity:"weekly",locked:false},
  {id:"P56",courseIdx:11,ins:"t11",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["2B"],students:20,day:"thu",period:1,room:"305",parity:"weekly",locked:false},
  {id:"P57",courseIdx:12,ins:"t12",type:"L",roomType:"online",freq:"weekly",phase:"full",cohorts:["2A","2B"],students:40,day:"wed",period:3,room:"ONLINE",parity:"weekly",locked:false},
  {id:"P58",courseIdx:12,ins:"t12",type:"S",roomType:"online",freq:"weekly",phase:"full",cohorts:["2A","2B"],students:40,day:"thu",period:3,room:"ONLINE",parity:"weekly",locked:false},
  {id:"P59",courseIdx:12,ins:"t12",type:"B",roomType:"online",freq:"weekly",phase:"full",cohorts:["2A","2B"],students:40,day:"fri",period:3,room:"ONLINE",parity:"weekly",locked:false},
  {id:"P60",courseIdx:15,ins:"t6",parallel:true,ins2:"t16",room2:"305",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["2A","2B"],students:40,day:"fri",period:4,room:"308",parity:"weekly",locked:false},
  {id:"P61",courseIdx:15,ins:"t6",parallel:true,ins2:"t16",room2:"305",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["2A","2B"],students:40,day:"wed",period:4,room:"308",parity:"weekly",locked:false},
  {id:"P62",courseIdx:15,ins:"t6",parallel:true,ins2:"t16",room2:"305",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["2A","2B"],students:40,day:"thu",period:4,room:"308",parity:"weekly",locked:false},
  {id:"P63",courseIdx:16,ins:"t5",type:"L",roomType:"lecture",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"thu",period:1,room:"303",parity:"weekly",locked:false},
  {id:"P64",courseIdx:16,ins:"t5",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"thu",period:2,room:"303",parity:"weekly",locked:false},
  {id:"P65",courseIdx:16,ins:"t5",type:"B",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"mon",period:1,room:"303",parity:"weekly",locked:false},
  {id:"P66",courseIdx:19,ins:"t6",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"wed",period:1,room:"303",parity:"weekly",locked:false},
  {id:"P67",courseIdx:19,ins:"t6",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"tue",period:2,room:"303",parity:"weekly",locked:false},
  {id:"P68",courseIdx:19,ins:"t6",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"fri",period:3,room:"303",parity:"weekly",locked:false},
  {id:"P69",courseIdx:20,ins:"t14",type:"L",roomType:"lecture",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"wed",period:2,room:"303",parity:"weekly",locked:false},
  {id:"P70",courseIdx:20,ins:"t14",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"wed",period:3,room:"303",parity:"weekly",locked:false},
  {id:"P71",courseIdx:20,ins:"t14",type:"B",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"thu",period:3,room:"303",parity:"weekly",locked:false},
  {id:"P72",courseIdx:4,ins:"t4",type:"L",roomType:"online",freq:"weekly",phase:"full",cohorts:["1A","1B"],students:48,day:"wed",period:4,room:"ONLINE",parity:"weekly",locked:false},
  {id:"P73",courseIdx:4,ins:"t4",type:"S",roomType:"online",freq:"weekly",phase:"full",cohorts:["1A","1B"],students:48,day:"mon",period:4,room:"ONLINE",parity:"weekly",locked:false},
  {id:"P74",courseIdx:5,ins:"t5",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1A"],students:24,day:"fri",period:4,room:"304",parity:"weekly",locked:false},
  {id:"P75",courseIdx:5,ins:"t5",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["1B"],students:24,day:"wed",period:2,room:"309",parity:"weekly",locked:false},
  {id:"P76",courseIdx:17,ins:"t8",type:"L",roomType:"lecture",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"tue",period:3,room:"303",parity:"weekly",locked:false},
  {id:"P77",courseIdx:17,ins:"t8",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"fri",period:4,room:"303",parity:"weekly",locked:false},
  {id:"P78",courseIdx:18,ins:"t5",type:"L",roomType:"lecture",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"wed",period:4,room:"303",parity:"weekly",locked:false},
  {id:"P79",courseIdx:18,ins:"t5",type:"S",roomType:"seminar",freq:"weekly",phase:"full",cohorts:["3A"],students:20,day:"thu",period:4,room:"303",parity:"weekly",locked:false},
  {id:"P80",courseIdx:25,ins:"t8",type:"S",roomType:"seminar",freq:"biweekly",phase:"h1",cohorts:["4A"],students:18,day:"wed",period:3,room:"307",parity:"odd",locked:false},
];

// ---------- Derivation ----------
const PHASES_OF = { full:["h1","h2"], h1:["h1"], h2:["h2"] };
const weeksOf = (len) => (len === "full" ? 16 : 8);
function deriveComponents(c) {
  const out = []; const w = weeksOf(c.length); const on = c.online;
  const add = (type, credits, room, combine=false) => {
    if (!credits) return;
    const perCredit = type === "L" ? 8 : 16;
    const perWeek = (credits * perCredit) / w;
    const weekly = Math.floor(perWeek + 1e-9); const rem = perWeek - weekly;
    const rt = on ? "online" : room;
    for (let i=0;i<weekly;i++) out.push({ type, freq:"weekly", room:rt, combine });
    if (rem >= 0.5 - 1e-9) out.push({ type, freq:"biweekly", room:rt, combine });
    else if (weekly === 0 && rem > 1e-9) out.push({ type, freq:"biweekly", room:rt, combine });
  };
  add("L", c.lec, "lecture", c.combineLec);
  add("S", c.sem, "seminar"); add("B", c.bonus, "seminar");
  add("Lab", c.lab, "lab"); add("P", c.prac, "practical");
  return out;
}
function expandSessions(courses) {
  const out = []; let n = 0;
  for (const c of courses) for (const comp of deriveComponents(c)) {
    const base = { courseIdx:c.idx, ins:c.t, type:comp.type, freq:comp.freq, phase:c.length, roomType:comp.room };
    // Parallel seminar: one event covering all cohorts, split into 2 level sections (2 rooms, 2 teachers)
    if (comp.type === "S" && c.parallel) { out.push(mk(n++, { ...base, parallel:true, ins2:c.t2||c.t }, c.cohorts)); continue; }
    const grouped = (comp.type === "L" && c.combineLec) || ((comp.type === "S" || comp.type === "B") && c.combineSem);
    if (grouped && c.cohorts.length > 1) out.push(mk(n++, base, c.cohorts));
    else for (const co of c.cohorts) out.push(mk(n++, base, [co]));
  }
  return out;
}
function mk(n, base, cohorts) {
  return { id:`S${n}`, ...base, cohorts, students:cohorts.reduce((s,id)=>s+cohort(id).students,0),
    day:null, period:null, room:null, parity:null, locked:false };
}

// ---------- Occupancy (semester-half × week-parity) ----------
const cellKey = (d,p,ph,pa) => `${d}|${p}|${ph}|${pa}`;
function occFactory(placed) {
  const occ = new Map();
  const get = (d,p,ph,pa) => { const k = cellKey(d,p,ph,pa); if (!occ.has(k)) occ.set(k, { ins:new Set(), rooms:new Set(), cohorts:new Set() }); return occ.get(k); };
  const mark = (s) => {
    const pars = s.parity === "weekly" ? ["odd","even"] : [s.parity];
    for (const ph of PHASES_OF[s.phase]) for (const pa of pars) {
      const c = get(s.day, s.period, ph, pa);
      c.ins.add(s.ins); if (s.ins2 && s.ins2!==s.ins) c.ins.add(s.ins2);
      if (s.room && s.room !== "ONLINE") c.rooms.add(s.room); if (s.room2 && s.room2 !== "ONLINE") c.rooms.add(s.room2);
      s.cohorts.forEach((x)=>c.cohorts.add(x));
    }
  };
  for (const s of placed) if (s.day) mark(s);
  return { get, mark };
}
function feasible(s, d, p, cp, get, seq, offMap) {
  if (offMap[s.ins] === d) return false;
  const pars = cp === "weekly" ? ["odd","even"] : [cp];
  for (const ph of PHASES_OF[s.phase]) for (const pa of pars) {
    const cell = get(d,p,ph,pa);
    if (cell.ins.has(s.ins)) return false;
    if (s.cohorts.some((c)=>cell.cohorts.has(c))) return false;
  }
  const si = slotIndex(d,p);
  if (s.type === "S" && seq.lec[s.courseIdx] != null && si <= seq.lec[s.courseIdx]) return false;
  if (s.type === "B" && seq.sem[s.courseIdx] != null && si <= seq.sem[s.courseIdx]) return false;
  return true;
}
// rooms that could legally hold this session at its current slot (free + fits), for manual override
function availableRoomsFor(session, placed, half) {
  if (session.roomType === "online" || session.parallel) return session.roomType==="online" ? ["ONLINE"] : null;
  const others = placed.filter((s)=>s.id!==session.id);
  const { get } = occFactory(others);
  const cp = session.parity; const pars = cp === "weekly" ? ["odd","even"] : [cp];
  const phases = half ? [half] : PHASES_OF[session.phase]; // only the half being edited
  const free = (rid) => { for (const ph of phases) for (const pa of pars) if (get(session.day, session.period, ph, pa).rooms.has(rid)) return false; return true; };
  // halls in priority order, then the rest; lab is fixed to 201
  const halls = hallsByPri();
  const others2 = ROOMS.filter((r)=>r.id!=="201" && !halls.includes(r.id)).map((r)=>r.id);
  const pool = session.roomType === "lab" ? ["201"] : [...halls, ...others2];
  const out = pool.filter((rid)=>{ const r = ROOMS.find((x)=>x.id===rid); return r && r.cap >= session.students && (rid===session.room || free(rid)); });
  if (session.room && session.room !== "ONLINE" && !out.includes(session.room)) out.unshift(session.room);
  return out;
}
function findRoom(s, d, p, cp, get, half) {
  const pars = cp === "weekly" ? ["odd","even"] : [cp];
  const phases = half ? [half] : PHASES_OF[s.phase];
  const free = (rid) => { for (const ph of phases) for (const pa of pars) if (get(d,p,ph,pa).rooms.has(rid)) return false; return true; };
  switch (s.roomType) {
    case "online": return "ONLINE";
    case "lab": { const r = ROOMS.find((x)=>x.id==="201"); return (r && r.cap >= s.students && free("201")) ? "201" : null; }
    default: {
      if (s.cohorts.length === 1) { const hr = homeRoom(s.cohorts[0]); const r = ROOMS.find((x)=>x.id===hr); if (hr && r.cap >= s.students && free(hr)) return hr; }
      for (const h of hallsByPri()) { const r = ROOMS.find((x)=>x.id===h); if (free(h) && r.cap >= s.students) return h; }
      return null;
    }
  }
}
const mulberry32 = (a) => () => { a|=0; a=a+0x6D2B79F5|0; let t=Math.imul(a^a>>>15,1|a); t=t+Math.imul(t^t>>>7,61|t)^t; return ((t^t>>>14)>>>0)/4294967296; };
const shuffle = (arr, rand) => { const a=[...arr]; for (let i=a.length-1;i>0;i--){ const j=Math.floor(rand()*(i+1)); [a[i],a[j]]=[a[j],a[i]]; } return a; };
function cost(s, d, p, cp, placed, prefMap, rand) {
  let c = 0;
  if (p === 4) c += 8; else if (p === 3) c += 1;
  if (prefMap[s.ins] === d) c -= 3;
  for (const co of s.cohorts) {
    const ds = placed.filter((x)=>x.day===d && x.cohorts.includes(co));
    c += ds.length * 2;
    const ps = ds.map((x)=>x.period).concat(p).sort((a,b)=>a-b);
    if (ps[ps.length-1]-ps[0]+1 > ps.length) c += 4;
  }
  if (cp !== "weekly") {
    const opp = cp === "odd" ? "even" : "odd";
    const pair = placed.some((x)=> x.day===d && x.period===p && x.parity===opp && x.cohorts.some((co)=>s.cohorts.includes(co)) && PHASES_OF[x.phase].some((ph)=>PHASES_OF[s.phase].includes(ph)));
    if (pair) c -= 7;
  }
  if (placed.filter((x)=>x.day===d && x.ins===s.ins).length >= 3) c += 6;
  c += rand() * 4; // exploration jitter → distinct candidates
  return c;
}
// ---------- incremental occupancy for the backtracking solver ----------
function makeOcc() {
  const occ = new Map();
  const cellOf = (d,p,ph,pa) => { const k = cellKey(d,p,ph,pa); let c = occ.get(k); if (!c) { c = { ins:{}, rooms:{}, coh:{} }; occ.set(k,c); } return c; };
  const bump = (o,k,v) => { o[k] = (o[k]||0)+v; if (!o[k]) delete o[k]; };
  const cells = (s) => { const out=[]; const pars = s.parity==="weekly"?["odd","even"]:[s.parity]; for (const ph of PHASES_OF[s.phase]) for (const pa of pars) out.push(cellOf(s.day,s.period,ph,pa)); return out; };
  const add = (s) => { for (const c of cells(s)) { bump(c.ins,s.ins,1); if (s.ins2 && s.ins2!==s.ins) bump(c.ins,s.ins2,1); if (s.room && s.room!=="ONLINE") bump(c.rooms,s.room,1); if (s.room2 && s.room2!=="ONLINE") bump(c.rooms,s.room2,1); s.cohorts.forEach((x)=>bump(c.coh,x,1)); } };
  const remove = (s) => { for (const c of cells(s)) { bump(c.ins,s.ins,-1); if (s.ins2 && s.ins2!==s.ins) bump(c.ins,s.ins2,-1); if (s.room && s.room!=="ONLINE") bump(c.rooms,s.room,-1); if (s.room2 && s.room2!=="ONLINE") bump(c.rooms,s.room2,-1); s.cohorts.forEach((x)=>bump(c.coh,x,-1)); } };
  const canPlace = (s,d,p,pa) => { const pars = pa==="weekly"?["odd","even"]:[pa]; for (const ph of PHASES_OF[s.phase]) for (const q of pars) { const c = cellOf(d,p,ph,q); if (c.ins[s.ins]) return false; if (s.ins2 && s.ins2!==s.ins && c.ins[s.ins2]) return false; for (const co of s.cohorts) if (c.coh[co]) return false; } return true; };
  const roomFree = (rid,s,d,p,pa) => { const pars = pa==="weekly"?["odd","even"]:[pa]; for (const ph of PHASES_OF[s.phase]) for (const q of pars) if (cellOf(d,p,ph,q).rooms[rid]) return false; return true; };
  return { add, remove, canPlace, roomFree };
}
// two seminar rooms for a parallel (level-split) session — prefer the cohorts' own home rooms
function pickParallelRooms(s,d,p,pa,occ) {
  if (s.roomType === "online") return ["ONLINE","ONLINE"];
  const perSec = Math.ceil(s.students/2);
  const seminar = ROOMS.filter((r)=>r.type==="seminar" && r.cap>=perSec).map((r)=>r.id);
  const homes = s.cohorts.map((c)=>homeRoom(c)).filter((h)=>h && seminar.includes(h));
  const ordered = [...new Set([...homes, ...seminar])];
  const free = ordered.filter((rid)=>occ.roomFree(rid,s,d,p,pa));
  return free.length>=2 ? [free[0], free[1]] : null;
}
function pickRoom(s,d,p,pa,occ) {
  if (s.roomType === "online") return "ONLINE";
  if (s.roomType === "lab") { const r = ROOMS.find((x)=>x.id==="201"); return (r && r.cap >= s.students && occ.roomFree("201",s,d,p,pa)) ? "201" : null; }
  // a single cohort that fits its own room stays there (3rd/4th-year lectures don't move to a hall)
  if (s.cohorts.length === 1) { const hr = homeRoom(s.cohorts[0]); const r = ROOMS.find((x)=>x.id===hr); if (hr && r.cap >= s.students && occ.roomFree(hr,s,d,p,pa)) return hr; }
  // combined groups (or anything too big for a home room) go to a lecture hall, in priority order
  for (const h of hallsByPri()) { const r = ROOMS.find((x)=>x.id===h); if (r.cap >= s.students && occ.roomFree(h,s,d,p,pa)) return h; }
  return null;
}
// ---------- Default soft-rule switches (see RULES_META) ----------
const DEFAULT_RULES = { compactDay:true, minGaps:true, balanceStudent:true, noSameCourseDay:true, biweeklyEdge:true, teacherPrefDay:true, teacherAvoid:true, pairBiweekly:true, balanceTeacher:true };
const RULES_HARD = [
  { en:"No student group is in two places at once", mn:"Нэг бүлэг зэрэг хоёр газар байхгүй" },
  { en:"No instructor is in two places at once", mn:"Нэг багш зэрэг хоёр газар байхгүй" },
  { en:"No room is double-booked", mn:"Нэг өрөө давхардахгүй" },
  { en:"A class never exceeds its room's capacity", mn:"Хичээл өрөөний багтаамжаас хэтрэхгүй" },
  { en:"Lecture is before its seminar and lab; seminar is before bonus seminar", mn:"Лекц нь семинар, лабораториос өмнө; семинар нь нэмэлт семинараас өмнө" },
  { en:"An instructor's fixed day-off is respected", mn:"Багшийн амрах өдрийг баримтална" },
  { en:"Parallel sections book both instructors and both rooms together", mn:"Зэрэгцээ хичээл хоёр багш, хоёр өрөөг зэрэг захиална" },
  { en:"Small single groups (3rd/4th year) keep lectures in their own room; halls only when a group is too big", mn:"Цөөн бүлэг (3,4-р курс) лекцээ өөрийн өрөөнд; том танхим зөвхөн шаардлагатай үед" },
];
const RULES_SOFT = [
  { id:"compactDay", en:"Keep each day compact — avoid the 13:50 late period and Tuesday afternoon", mn:"Өдрийг нягт байлгах — 13:50-ийн оройн цаг ба Мягмар үдээс хойшхоос зайлсхийх" },
  { id:"minGaps", en:"No mid-day gaps — keep each group's day a single unbroken block (strict when on)", mn:"Өдрийн дунд цонхгүй — бүлгийн өдрийг тасралтгүй байлгах (асаалттай үед хатуу)" },
  { id:"balanceStudent", en:"Balance a group's classes evenly across the week (e.g. 2·2·2·2·2, not 3·3·3·1·0)", mn:"Бүлгийн хичээлийг долоо хоногт жигд хуваарилах (ж: 2·2·2·2·2, 3·3·3·1·0 биш)" },
  { id:"noSameCourseDay", en:"No two of the same session type in a day — e.g. two seminars of a course; lecture+seminar and lecture+lab stay allowed (strict when on)", mn:"Нэг өдөрт ижил төрлийн хоёр хичээл байхгүй — ж: нэг хичээлийн хоёр семинар; лекц+семинар, лекц+лаб зөвшөөрөгдөнө (асаалттай үед хатуу)" },
  { id:"biweeklyEdge", en:"Put biweekly classes at the start or end of the day", mn:"2 долоо хоног тутмын хичээлийг өдрийн эхэнд/төгсгөлд" },
  { id:"teacherPrefDay", en:"Respect an instructor's preferred day", mn:"Багшийн дуртай өдрийг харгалзах" },
  { id:"teacherAvoid", en:"Honour instructor time requests — a period they'd rather not teach (e.g. no 1st period for school pickup)", mn:"Багшийн цагийн хүсэлтийг харгалзах — заахыг хүсэхгүй цаг (ж: хүүхэд авах тул 1-р цаггүй)" },
  { id:"pairBiweekly", en:"Pair odd/even biweekly classes in the same slot", mn:"Сондгой/тэгш 7 хоногийн хичээлийг нэг нүдэнд хослуулах" },
  { id:"balanceTeacher", en:"Balance instructor load across the week", mn:"Багшийн ачааллыг долоо хоногт тэнцвэржүүлэх" },
];
// ---------- backtracking scheduler (hard rules always on; soft rules from `rules`) ----------
function scheduleOnce(courses, teachers, locked, rand, rules = DEFAULT_RULES) {
  const offMap = {}, prefMap = {}, avoidMap = {}, avoidPer = {};
  teachers.forEach((t)=>{ offMap[t.id]=t.off; prefMap[t.id]=t.pref;
    avoidMap[t.id] = new Set((t.avoid||[]).filter((a)=>a.day!=="any").map((a)=>`${a.day}|${a.period}`));
    avoidPer[t.id] = new Set((t.avoid||[]).filter((a)=>a.day==="any").map((a)=>a.period)); });
  const all = expandSessions(courses);
  // per-cohort, per-half target classes/day (total ÷ 5 teaching days) → drives even distribution
  const halfTotal = {};
  for (const s of all) for (const co of s.cohorts) for (const h of PHASES_OF[s.phase]) { (halfTotal[co] = halfTotal[co] || { h1:0, h2:0 })[h]++; }
  const dayTarget = {}; for (const co in halfTotal) dayTarget[co] = { h1:Math.ceil(halfTotal[co].h1/5), h2:Math.ceil(halfTotal[co].h2/5) };
  const lockKey = (s)=>`${s.courseIdx}|${s.type}|${s.freq}|${s.phase}|${s.cohorts.join("+")}`;
  const byKey = {};
  locked.forEach((ls)=>{ (byKey[lockKey(ls)] = byKey[lockKey(ls)] || []).push(ls); });
  const occ = makeOcc(); const placed = []; const pending = [];
  for (const s of all) {
    const k = lockKey(s);
    if (byKey[k] && byKey[k].length) { const ls = byKey[k].shift(); const ps = { ...s, day:ls.day, period:ls.period, room:ls.room, parity:ls.parity, locked:true }; placed.push(ps); occ.add(ps); }
    else pending.push(s);
  }
  // place each course's sessions together (lecture, then seminar/lab, then bonus) so ordering stays local
  const cCount = {}; pending.forEach((s)=>cCount[s.courseIdx]=(cCount[s.courseIdx]||0)+1);
  const rp = (rt)=> rt==="lab"?0:rt==="lecture"?1:3;
  const order = pending.sort((a,b)=>
    (cCount[b.courseIdx]-cCount[a.courseIdx]) || (a.courseIdx-b.courseIdx) ||
    COMP_PHASE[a.type]-COMP_PHASE[b.type] || rp(a.roomType)-rp(b.roomType));
  const dayOrder = shuffle(DAYS, rand);
  // HARD: lecture before seminar & lab; seminar before bonus (positions in the Tue→Mon week)
  const maxPos = (ci,type)=>{ let m=-1; for (const x of placed) if (x.courseIdx===ci && x.type===type) { const q=slotIndex(x.day,x.period); if (q>m) m=q; } return m; };
  const seqOK = (s,d,p)=>{ const pos=slotIndex(d,p);
    if (s.type==="S" || s.type==="Lab") { const L=maxPos(s.courseIdx,"L"); if (L>=0 && pos<=L) return false; }
    if (s.type==="B") { const S=maxPos(s.courseIdx,"S"); if (S>=0 && pos<=S) return false; const L=maxPos(s.courseIdx,"L"); if (L>=0 && pos<=L) return false; }
    return true; };
  const placeCost = (s,d,p,pa) => {
    let c = 0;
    if (rules.compactDay) { if (p === 4) c += 60; else if (p === 3) c += 1; } // avoid the late 13:50 period (soft)
    if (rules.teacherPrefDay && prefMap[s.ins] === d) c -= 3;
    if (rules.teacherAvoid && ((avoidMap[s.ins] && avoidMap[s.ins].has(`${d}|${p}`)) || (avoidPer[s.ins] && avoidPer[s.ins].has(p)))) c += 40;
    if (rules.minGaps) for (const co of s.cohorts) for (const h of PHASES_OF[s.phase]) {
      const per=new Set(); for (const x of placed) if (x.day===d && x.cohorts.includes(co) && PHASES_OF[x.phase].includes(h)) per.add(x.period);
      if (per.size) { const arr=[...per], mn=Math.min(...arr), mx=Math.max(...arr);
        const gapBefore=(mx-mn+1)-per.size; const ns=per.has(p)?per.size:per.size+1;
        const gapAfter=(Math.max(mx,p)-Math.min(mn,p)+1)-ns;
        c += (gapAfter-gapBefore)*80; }        // strongly penalize opening a mid-day gap (per half)
      else c += (p-1)*3;                        // first class of the day: prefer starting at P1
    }
    if (rules.balanceStudent) for (const co of s.cohorts) for (const h of PHASES_OF[s.phase]) {
      let n=0; for (const x of placed) if (x.day===d && x.cohorts.includes(co) && PHASES_OF[x.phase].includes(h)) n++;
      const tgt = (dayTarget[co] && dayTarget[co][h]) || 3;
      c += n*2 + (n>=tgt ? 24 : 0); // fill days up to the target evenly before overloading any one day
    }
    if (rules.noSameCourseDay) { for (const x of placed) if (x.courseIdx===s.courseIdx && x.type===s.type && x.day===d && x.period!==p && x.cohorts.some((co)=>s.cohorts.includes(co))) { c += 200; break; } } // strongly avoid the same course+type at another period the same day
    if (pa !== "weekly") {
      if (rules.pairBiweekly) { const opp = pa==="odd"?"even":"odd"; for (const x of placed) if (x.day===d && x.period===p && x.parity===opp && x.cohorts.some((co)=>s.cohorts.includes(co))) { c -= 6; break; } }
      if (rules.biweeklyEdge) { // biweekly must be first/last of the day, else the off-week has a mid-day gap
        let mn=99,mx=0,any=false; for (const x of placed) if (x.day===d && x.cohorts.some((co)=>s.cohorts.includes(co))) { any=true; if (x.period<mn) mn=x.period; if (x.period>mx) mx=x.period; }
        c += (p-1)*22;                             // pull toward P1 (an edge, and lecture-before-seminar friendly)
        if (any && p>mn && p<mx) c += 45;          // heavy penalty for landing in the interior of the block
      }
    }
    if (rules.balanceTeacher) { let n=0; for (const x of placed) if (x.day===d && x.ins===s.ins) n++; if (n>=3) c += 6; }
    c += rand() * 3;
    return c;
  };
  const contiguousOK = (s,d,p,except) => { // each half a student attends must be an unbroken block
    for (const half of PHASES_OF[s.phase]) for (const co of s.cohorts) {
      let mn=99, mx=0, any=false, has=false;
      for (const x of placed) if (x!==except && x.day===d && x.cohorts.includes(co) && PHASES_OF[x.phase].includes(half)) {
        any=true; if (x.period<mn) mn=x.period; if (x.period>mx) mx=x.period; if (x.period===p) has=true;
      }
      if (!any || has) continue;                 // first class of this half's day, or same slot → fine
      if (p !== mn-1 && p !== mx+1) return false; // otherwise must extend the block — no gap in this half
    }
    return true;
  };
  let ALLOWP4 = true;   // single pass: P4 is allowed but heavily cost-penalized, so it's only used when needed
  const domain = (s, relax) => {
    relax = relax || {};
    const opts = []; const pars = s.freq==="weekly" ? ["weekly"] : ["odd","even"];
    for (const day of dayOrder) for (const per of PERIODS) for (const pa of pars) {
      if (!ALLOWP4 && per.id === 4) continue;
      if (day.id === "tue" && per.id >= TUE_CUTOFF_PERIOD) continue; // Tuesday afternoon: hard block (no classes after 13:45)
      if (offMap[s.ins] === day.id) continue;
      if (!seqOK(s, day.id, per.id)) continue;
      if (!relax.gaps && rules.minGaps && !contiguousOK(s, day.id, per.id)) continue; // per-half contiguity (no mid-day gaps)
      if (!relax.sameType && rules.noSameCourseDay && placed.some((x)=>x.courseIdx===s.courseIdx && x.type===s.type && x.day===day.id && x.period!==per.id && x.cohorts.some((co)=>s.cohorts.includes(co)))) continue; // avoid same course+type twice in a day (lecture+seminar/lecture+lab still allowed)
      if (!occ.canPlace(s, day.id, per.id, pa)) continue;
      let room, room2=null;
      if (s.parallel) { const rr = pickParallelRooms(s, day.id, per.id, pa, occ); if (!rr) continue; [room, room2] = rr; }
      else { room = pickRoom(s, day.id, per.id, pa, occ); if (!room) continue; }
      opts.push({ d:day.id, p:per.id, pa, room, room2, c:placeCost(s, day.id, per.id, pa) });
    }
    opts.sort((a,b)=>a.c-b.c);
    return opts;
  };
  // strict first; if a class can't fit, allow a mid-day gap (never a same-type repeat) so we still place it
  const domainWithFallback = (s) => {
    let o = domain(s);
    if (!o.length && rules.minGaps) o = domain(s, { gaps:true }); // allow a gap to fit it — same-type-per-day stays hard
    return o;
  };
  let nodes = 0; const BUDGET = 12000;
  let bestPlaced = placed.slice(); let bestCount = placed.length; // best partial seen (starts at locked-only)
  function dfs(i, relax) {
    if (placed.length > bestCount) { bestCount = placed.length; bestPlaced = placed.slice(); } // remember deepest partial
    if (i === order.length) return true;
    if (++nodes > BUDGET) return false;
    const s = order[i];
    for (const o of (relax ? domainWithFallback(s) : domain(s))) {
      const ps = { ...s, day:o.d, period:o.p, parity:o.pa, room:o.room, room2:o.room2 };
      occ.add(ps); placed.push(ps);
      if (dfs(i+1, relax)) return true;
      occ.remove(ps); placed.pop();
    }
    return false;
  }
  let ok = dfs(0, false);                     // strict pass: fully honor no-gaps + no-same-type
  if (!ok) { nodes = 0; ok = dfs(0, true); }  // only if strict can't place everything, relax those soft rules just enough
  const finalPlaced = ok ? placed : bestPlaced;   // never return empty — fall back to the best partial found
  if (rules.noSameCourseDay) { // repair: spread same-course sessions onto separate days now that all blocks exist
    const rocc = makeOcc(); finalPlaced.forEach((s)=>rocc.add(s));
    const inList = (x)=>finalPlaced.indexOf(x)>=0;
    const contigOK = (s,d,p,except)=>{ for (const half of PHASES_OF[s.phase]) for (const co of s.cohorts) {
        let mn=99,mx=0,any=false,has=false;
        for (const x of finalPlaced) if (x!==except && x.day===d && x.cohorts.includes(co) && PHASES_OF[x.phase].includes(half)) { any=true; if (x.period<mn)mn=x.period; if (x.period>mx)mx=x.period; if (x.period===p)has=true; }
        if (!any||has) continue; if (p!==mn-1 && p!==mx+1) return false; } return true; };
    const dayPeriods = (co,d,half)=>{ const ps=[]; for (const x of finalPlaced) if (x.day===d && x.cohorts.includes(co) && PHASES_OF[x.phase].includes(half)) ps.push(x.period); return ps; };
    const removalSafe = (s)=>{ for (const co of s.cohorts) for (const half of PHASES_OF[s.phase]) { const ps=dayPeriods(co,s.day,half); if (ps.length>1 && s.period!==Math.min(...ps) && s.period!==Math.max(...ps)) return false; } return true; };
    for (let pass=0; pass<4; pass++) { let changed=false;
      for (const s of finalPlaced) {
        if (s.locked || s.parity!=="weekly") continue;
        const twin = finalPlaced.find((x)=>x!==s && x.courseIdx===s.courseIdx && x.type===s.type && x.day===s.day && x.period!==s.period && x.cohorts.some((co)=>s.cohorts.includes(co)));
        if (!twin || !removalSafe(s)) continue;
        rocc.remove(s); let moved=false;
        for (const day of dayOrder) {
          if (day.id===s.day || offMap[s.ins]===day.id) continue;
          if (finalPlaced.some((x)=>x!==s && x.courseIdx===s.courseIdx && x.type===s.type && x.day===day.id && x.cohorts.some((co)=>s.cohorts.includes(co)))) continue;
          for (const per of PERIODS) {
            if (day.id==="tue" && per.id>=TUE_CUTOFF_PERIOD) continue; // hard block: no classes Tue after 13:45 (matches domain() and compaction)
            if (!seqOK(s, day.id, per.id)) continue;
            if (rules.minGaps && !contigOK(s, day.id, per.id, s)) continue;
            if (!rocc.canPlace(s, day.id, per.id, s.parity)) continue;
            const room = s.parallel ? pickParallelRooms(s,day.id,per.id,s.parity,rocc) : pickRoom(s,day.id,per.id,s.parity,rocc);
            if (!room) continue;
            s.day=day.id; s.period=per.id; if (s.parallel){ s.room=room[0]; s.room2=room[1]; } else s.room=room;
            rocc.add(s); moved=true; changed=true; break;
          }
          if (moved) break;
        }
        if (!moved) rocc.add(s);
      }
      if (!changed) break;
    }
  }
  if (rules.minGaps) { // compaction: slide classes into free slots so each group's day is contiguous (no mid-day gaps)
    const gocc = makeOcc(); finalPlaced.forEach((s)=>gocc.add(s));
    const bucketGaps = (buckets) => { let g=0; for (const [co,d,h] of buckets) { const ps=[]; for (const x of finalPlaced) if (x.day===d && x.cohorts.includes(co) && PHASES_OF[x.phase].includes(h)) ps.push(x.period); if (ps.length) { const u=[...new Set(ps)]; g += (Math.max(...u)-Math.min(...u)+1) - u.length; } } return g; };
    const sameTypeClash = (s,d,p) => finalPlaced.some((x)=>x!==s && x.courseIdx===s.courseIdx && x.type===s.type && x.day===d && x.period!==p && x.cohorts.some((co)=>s.cohorts.includes(co)));
    for (let pass=0; pass<8; pass++) { let improved=false;
      for (const s of finalPlaced) {
        if (s.locked) continue;
        const od=s.day, op=s.period, or=s.room, or2=s.room2;
        gocc.remove(s);
        let best=null, bestDelta=0, bestTie=Infinity;
        for (const day of dayOrder) for (const per of PERIODS) {
          if (day.id===od && per.id===op) continue;
          if (day.id==="tue" && per.id>=TUE_CUTOFF_PERIOD) continue;
          if (offMap[s.ins]===day.id) continue;
          if (sameTypeClash(s, day.id, per.id)) continue;
          if (!seqOK(s, day.id, per.id)) continue;
          if (!gocc.canPlace(s, day.id, per.id, s.parity)) continue;
          const room = s.parallel ? pickParallelRooms(s,day.id,per.id,s.parity,gocc) : pickRoom(s,day.id,per.id,s.parity,gocc);
          if (!room) continue;
          const buckets=[]; for (const co of s.cohorts) for (const dd of [od, day.id]) for (const h of PHASES_OF[s.phase]) buckets.push([co,dd,h]);
          const base = bucketGaps(buckets);
          s.day=day.id; s.period=per.id;
          const cand = bucketGaps(buckets);
          s.day=od; s.period=op;
          const delta = cand - base;
          const tie = per.id + (per.id===4?5:0);           // among equal-gap moves, prefer earlier / avoid P4
          if (delta < bestDelta - 1e-9 || (Math.abs(delta-bestDelta)<1e-9 && delta<0 && tie<bestTie)) { bestDelta=delta; bestTie=tie; best={ day:day.id, period:per.id, room }; }
        }
        if (best) { s.day=best.day; s.period=best.period; if (s.parallel){ s.room=best.room[0]; s.room2=best.room[1]; } else s.room=best.room; gocc.add(s); improved=true; }
        else { s.day=od; s.period=op; s.room=or; s.room2=or2; gocc.add(s); }
      }
      if (!improved) break;
    }
  }
  const ids = new Set(finalPlaced.map((s)=>s.id));
  const unplaced = order.filter((s)=>!ids.has(s.id));
  return { placed: finalPlaced, unplaced };
}
function rankCandidates(cands) {
  cands.sort((x,y)=>
    (x.unplaced.length - y.unplaced.length) ||          // 1. hard constraints (all placed)
    (x.score.metrics.dupes - y.score.metrics.dupes) ||  // 2. fewest same-course-per-day repeats
    (x.score.metrics.gaps  - y.score.metrics.gaps)  ||  // 3. fewest mid-day breaks
    (x.score.metrics.intBW - y.score.metrics.intBW) ||  // 4. biweekly classes at the day edge
    (x.score.metrics.imb   - y.score.metrics.imb)   ||  // 5. best day-to-day distribution
    (y.score.overall - x.score.overall));               // 6. remaining soft preferences
  return cands;
}
function generateCandidates(courses, teachers, locked, attempts=8, rules=DEFAULT_RULES, seedBase=0) {
  const cands = []; const seen = new Set(); const mix = seedBase>>>0; // seedBase 0 reproduces the original fixed stream; a per-Generate base re-rolls
  for (let a=0;a<attempts;a++) {
    const res = scheduleOnce(courses, teachers, locked, mulberry32((((a*2654435761 + 9973) >>> 0) ^ mix) >>> 0), rules);
    const sig = res.placed.map((s)=>`${s.courseIdx}${s.type}${s.cohorts.join("")}${s.day}${s.period}${s.parity}`).sort().join("|");
    if (seen.has(sig)) continue; seen.add(sig);
    cands.push({ placed:res.placed, unplaced:res.unplaced, score:scoreSchedule(res.placed, res.unplaced, teachers) });
  }
  return rankCandidates(cands);
}
// UI version: yields to the browser between attempts so the page stays responsive
async function generateCandidatesAsync(courses, teachers, locked, attempts, rules, onProgress, seedBase=0) {
  const cands = []; const seen = new Set(); const mix = seedBase>>>0;
  for (let a=0;a<attempts;a++) {
    const res = scheduleOnce(courses, teachers, locked, mulberry32((((a*2654435761 + 9973) >>> 0) ^ mix) >>> 0), rules);
    const sig = res.placed.map((s)=>`${s.courseIdx}${s.type}${s.cohorts.join("")}${s.day}${s.period}${s.parity}`).sort().join("|");
    if (!seen.has(sig)) { seen.add(sig); cands.push({ placed:res.placed, unplaced:res.unplaced, score:scoreSchedule(res.placed, res.unplaced, teachers) }); }
    if (onProgress) onProgress(a+1, attempts);
    await new Promise((r)=>setTimeout(r, 0)); // let the UI breathe between attempts
  }
  return rankCandidates(cands);
}


// ---------- Capacity check (over-subscription) ----------
// Available student slots per half: Tue P1–P3 (3) + Wed/Thu/Fri/Mon P1–P4 (16) = 19
const SLOTS_PER_HALF = DAYS.length*PERIODS.length - (PERIODS.length - (TUE_CUTOFF_PERIOD-1)); // Tuesday afternoon blocked
function cohortLoad(courses) {
  const load = {}; COHORTS.forEach((c)=>load[c.id]={ h1:0, h2:0 });
  for (const s of expandSessions(courses)) {
    const w = s.freq==="weekly" ? 1 : 0.5;     // a biweekly session can share a slot with its pair
    for (const co of s.cohorts) for (const half of PHASES_OF[s.phase]) if (load[co]) load[co][half]+=w;
  }
  return load; // {cohortId:{h1,h2}} — slot demand per half; compare to SLOTS_PER_HALF
}

// ---------- Collision detection (a group/instructor/room in two places the same week) ----------
function collisionCells(placed, half=null) {
  const inHalf = (s)=> half ? PHASES_OF[s.phase].includes(half) : true;
  const list = placed.filter(inHalf);
  const parsOf = (s)=> s.parity==="weekly" ? ["odd","even"] : [s.parity];
  const halvesOf = (s)=> half ? [half] : PHASES_OF[s.phase];
  const buckets = {};
  for (const s of list) for (const h of halvesOf(s)) for (const pa of parsOf(s)) {
    const ents = [];
    s.cohorts.forEach((co)=>ents.push(["group", co]));
    [s.ins, s.ins2].filter(Boolean).forEach((ti)=>ents.push(["instructor", ti]));
    [s.room, s.room2].filter((r)=>r && r!=="ONLINE").forEach((r)=>ents.push(["room", r]));
    for (const [kind, ent] of ents) { const k=`${kind}|${ent}|${s.day}|${s.period}|${h}|${pa}`; (buckets[k]=buckets[k]||new Set()).add(s); }
  }
  const seen = new Set(); const out = [];
  for (const k in buckets) { const v=[...buckets[k]]; if (v.length>1) { const [kind,ent,day,period]=k.split("|");
    const dk=`${kind}|${ent}|${day}|${period}`; if (seen.has(dk)) continue; seen.add(dk);
    out.push({ kind, ent, day, period:+period, sessions:v }); } }
  return out;
}

// ---------- Scoring ----------
// half = "h1" | "h2" scores just that half; null/undefined scores the whole year (both halves)
function scoreSchedule(placed, unplaced, teachers, half=null) {
  const halves = half ? [half] : ["h1","h2"];
  const inHalf = (s) => half ? PHASES_OF[s.phase].includes(half) : true;
  placed = placed.filter(inHalf); unplaced = unplaced.filter(inHalf);
  let gaps=0, imb=0, p4=0, tueLate=0, intBW=0;
  for (const h of halves) {
    const cd = {}; COHORTS.forEach((c)=>{ cd[c.id]={}; DAYS.forEach((d)=>cd[c.id][d.id]=new Set()); });
    placed.forEach((s)=>{ if (PHASES_OF[s.phase].includes(h)) s.cohorts.forEach((c)=>{ if (cd[c]) cd[c][s.day].add(s.period); }); });
    for (const c of COHORTS) { const pd=[];
      for (const d of DAYS) { const arr=[...cd[c.id][d.id]].sort((a,b)=>a-b); pd.push(arr.length);
        if (arr.length) gaps += (arr[arr.length-1]-arr[0]+1)-arr.length; }
      if (pd.some((x)=>x>0)) imb += Math.max(...pd)-Math.min(...pd); // spread across ALL 5 days: cramming into fewer days is imbalance, not balance
    }
    placed.forEach((s)=>{ if (s.parity!=="weekly" && PHASES_OF[s.phase].includes(h)) s.cohorts.forEach((co)=>{ if (!cd[co]) return; const arr=[...cd[co][s.day]]; if (arr.length){ const mn=Math.min(...arr), mx=Math.max(...arr); if (s.period>mn && s.period<mx) intBW++; } }); });
  }
  placed.forEach((s)=>{ if (s.period===4) p4++; if (s.day==="tue" && s.period>=4) tueLate++; });
  let dupes=0; { const seen={}; placed.forEach((s)=>s.cohorts.forEach((c)=>{ const k=c+"|"+s.day+"|"+s.courseIdx+"|"+s.type; (seen[k]=seen[k]||new Set()).add(s.period); }));
    for (const k in seen) if (seen[k].size>1) dupes += seen[k].size-1; } // same course + same type, same day (e.g. two seminars)
  let stud = 100 - gaps*6 - p4*2 - tueLate*15 - dupes*3;
  let balance = 100 - imb*2; // k=2: with the corrected all-days spread (~15-26), realistic schedules land ~50-70 instead of clamping to 0
  let instr=100, over=0, hits=0, tot=0;
  for (const ins of teachers) {
    const byDay={}; placed.filter((s)=>s.ins===ins.id || s.ins2===ins.id).forEach((s)=>byDay[s.day]=(byDay[s.day]||0)+1);
    Object.values(byDay).forEach((n)=>{if(n>=4)over++;});
    if (ins.pref) { tot++; if (byDay[ins.pref]) hits++; }
  }
  instr -= over*8; if (tot) instr -= Math.round((1-hits/tot)*10);
  const combined = placed.filter((s)=>s.type==="L"&&s.cohorts.length>1).length;
  const room = Math.min(100, 85 + combined);
  const collisions = collisionCells(placed, half).length;
  const hard = (unplaced.length===0 && collisions===0) ? 100 : Math.max(0, 100 - unplaced.length*12 - collisions*10);
  const cl = (x)=>Math.max(0,Math.min(100,Math.round(x)));
  const parts = { hard:cl(hard), student:cl(stud), instructor:cl(instr), balance:cl(balance), room:cl(room) };
  const overall = cl(parts.hard*0.4+parts.student*0.30+parts.balance*0.15+parts.instructor*0.10+parts.room*0.05);
  return { parts, overall, metrics:{ gaps, p4, tueLate, imb, intBW, dupes, collisions } };
}

// ---------- Teacher schedule rating ----------
function teacherRating(id, placed) {
  const ss = placed.filter((s)=>s.ins===id || s.ins2===id);
  const byDay = {}; ss.forEach((s)=>{ (byDay[s.day]=byDay[s.day]||[]).push(s.period); });
  let overload=0, gaps=0, longRuns=0, maxRun=0;
  const days = Object.keys(byDay);
  for (const d of days) {
    const ps = [...new Set(byDay[d])].sort((a,b)=>a-b);
    if (ps.length >= 4) overload++;
    gaps += (ps[ps.length-1]-ps[0]+1) - ps.length;
    let run=1, best=1;
    for (let i=1;i<ps.length;i++){ if (ps[i]===ps[i-1]+1){ run++; best=Math.max(best,run);} else run=1; }
    maxRun = Math.max(maxRun, best);
    if (best>=4) longRuns+=2; else if (best===3) longRuns+=1;
  }
  const cl = (x)=>Math.max(0,Math.min(100,Math.round(x)));
  const workload = cl(100 - overload*22 - Math.max(0, ss.length-12)*4);
  const breaks = cl(100 - gaps*14);
  const consecutive = cl(100 - longRuns*16 - (maxRun>=4?10:0));
  const overall = cl(100 - overload*15 - gaps*7 - longRuns*8);
  const perHalf = { h1:0, h2:0 };
  ss.forEach((s)=>{ if (PHASES_OF[s.phase].includes("h1")) perHalf.h1++; if (PHASES_OF[s.phase].includes("h2")) perHalf.h2++; });
  const courseCount = new Set(ss.map((s)=>s.courseIdx)).size;
  return { overall, workload, breaks, consecutive, total:ss.length, perHalf, courseCount, days:days.length, gaps, maxRun, overload };
}

// ---------- Move validation ----------
function validateMove(session, d, p, placed, lang, teachers, half) {
  const offMap = {}; teachers.forEach((t)=>offMap[t.id]=t.off);
  const others = placed.filter((s)=>s.id!==session.id);
  const { get } = occFactory(others);
  const reasons = [];
  const cp = session.parity; const pars = cp === "weekly" ? ["odd","even"] : [cp];
  const phases = half ? [half] : PHASES_OF[session.phase]; // only check the half being edited
  const dName = dayLabel(lang, DAYS.find((x)=>x.id===d));
  if (d==="tue" && p>=TUE_CUTOFF_PERIOD) reasons.push({level:"conflict", text:tr(lang,"tueBlocked")});
  if (offMap[session.ins]===d) reasons.push({level:"conflict", text:`${teacherName(teachers,session.ins)} ${tr(lang,"isOffOn")} ${dName}`});
  let insHit=false, cohHit=false;
  for (const ph of phases) for (const pa of pars) { const cell = get(d,p,ph,pa);
    if (cell.ins.has(session.ins) || (session.ins2 && session.ins2!==session.ins && cell.ins.has(session.ins2))) insHit=true;
    if (session.cohorts.some((c)=>cell.cohorts.has(c))) cohHit=true; }
  if (insHit) reasons.push({level:"conflict", text:`${teacherName(teachers,session.ins)} — ${tr(lang,"alreadyBooked")}`});
  if (cohHit) reasons.push({level:"conflict", text:`${session.cohorts.join("+")} — ${tr(lang,"alreadyBooked")}`});
  let room, room2=null;
  if (session.parallel) {
    const perSec = Math.ceil(session.students/2);
    const free = (rid)=>{ for (const ph of phases) for (const pa of pars) if (get(d,p,ph,pa).rooms.has(rid)) return false; return true; };
    const homes = session.cohorts.map((c)=>homeRoom(c)).filter(Boolean);
    const pool = [...new Set([...homes, ...ROOMS.filter((r)=>r.type==="seminar").map((r)=>r.id)])].filter((rid)=>{ const r=ROOMS.find((x)=>x.id===rid); return r && r.cap>=perSec && free(rid); });
    if (session.roomType==="online") { room="ONLINE"; room2="ONLINE"; }
    else if (pool.length>=2) { room=pool[0]; room2=pool[1]; }
    if (!room) reasons.push({level:"conflict", text:tr(lang,"noRoomFree")});
  } else {
    room = findRoom(session, d, p, cp, get, half);
    if (!room) reasons.push({level:"conflict", text:tr(lang,"noRoomFree")});
  }
  if (p===4 && !reasons.length) reasons.push({level:"warning", text:`${tr(lang,"p4used")}`});
  if (!reasons.some((r)=>r.level==="conflict")) { const clash = others.some((x)=>x.courseIdx===session.courseIdx && x.type===session.type && x.day===d && x.cohorts.some((c)=>session.cohorts.includes(c))); if (clash) reasons.push({level:"warning", text:tr(lang,"sameTypeDay")}); }
  const status = reasons.some((r)=>r.level==="conflict")?"conflict":reasons.some((r)=>r.level==="warning")?"warning":"valid";
  return { status, reasons, room, room2 };
}

// ---------- helpers ----------
const teacherName = (teachers, id) => teachers.find((t)=>t.id===id)?.name ?? id;
const courseName = (courses, idx) => courses.find((c)=>c.idx===idx)?.name ?? defClassName(idx);
const phaseVisible = (phase, semView) => phase === "full" || phase === semView;
// lock is tracked per half (lockH1/lockH2); old single `locked` falls back to both halves
const lockedIn = (s, half) => !!(half==="h1" ? (s.lockH1 ?? s.locked) : (s.lockH2 ?? s.locked));
const clone = (x) => JSON.parse(JSON.stringify(x));

function exportCSV(placed, lang, courses, teachers, semView) {
  const rows = [[tr(lang,"day"), tr(lang,"period"), ...COHORTS.map((c)=>c.id)]];
  for (const d of DAYS) for (const p of PERIODS) {
    const row = [dayLabel(lang,d), p.label];
    for (const c of COHORTS) {
      const list = placed.filter((x)=>x.day===d.id && x.period===p.id && x.cohorts.includes(c.id) && phaseVisible(x.phase, semView));
      row.push(list.map((s)=>{ const par = s.parity==="weekly" ? "" : ` [${s.parity==="odd"?tr(lang,"oddWk"):tr(lang,"evenWk")}]`;
        return `${courseName(courses,s.courseIdx)} /${compTag(lang,s.type)}/ ${teacherName(teachers,s.ins)} /${s.room}/${par}`; }).join(" | "));
    }
    rows.push(row);
  }
  const csv = rows.map((r)=>r.map((x)=>`"${String(x).replace(/"/g,'""')}"`).join(",")).join("\n");
  const blob = new Blob(["\uFEFF"+csv], { type:"text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href=url; a.download="timetable.csv"; a.click(); URL.revokeObjectURL(url);
}

// ============================================================================
//  Course-data workbook  (round-trip: app ⇄ Excel template)
// ============================================================================
const LEN_TO_LABEL = { full:"Full semester", h1:"First half", h2:"Second half" };
const labelToLen = (s) => { const v=String(s||"").toLowerCase();
  if (v.includes("first")||v.includes("h1")||v==="1") return "h1";
  if (v.includes("second")||v.includes("h2")||v==="2") return "h2"; return "full"; };
const yes = (v) => { const s=String(v??"").trim().toLowerCase(); return s==="yes"||s==="true"||s==="1"||s==="y"||s==="тийм"; };

function buildCourseWorkbook(courses, teachers, termStart, roomPri) {
  const tname = (id)=> (teachers.find((t)=>t.id===id)||{}).name || id;
  const cHead = ["Course name","Instructor","Cohorts","Length","Lecture","Seminar","Bonus","Lab","Practical","Shared lecture","Shared seminar","Online","Parallel seminar","Teacher 2"];
  const cRows = [cHead, ...courses.map((c)=>[
    c.name, tname(c.t), c.cohorts.join(","), LEN_TO_LABEL[c.length]||"Full semester",
    c.lec||0, c.sem||0, c.bonus||0, c.lab||0, c.prac||0,
    c.combineLec?"Yes":"No", c.combineSem?"Yes":"No", c.online?"Yes":"No",
    c.parallel?"Yes":"No", c.parallel?tname(c.t2):"",
  ])];
  const iHead = ["Instructor","Fixed day off","Avoid P1","Avoid P2","Avoid P3","Avoid P4"];
  const dayName = { mon:"Monday",tue:"Tuesday",wed:"Wednesday",thu:"Thursday",fri:"Friday" };
  const iRows = [iHead, ...teachers.map((t)=>{ const av=new Set((t.avoid||[]).filter((a)=>a.day==="any").map((a)=>a.period));
    return [t.name, t.off?dayName[t.off]:"", av.has(1)?"Yes":"", av.has(2)?"Yes":"", av.has(3)?"Yes":"", av.has(4)?"Yes":""]; })];
  const sRows = [["Setting","Value"],["Academic year","2026-2027"],["Semester","Fall"],["Start date",termStart||DEFAULT_TERM_START],
    ["",""],["Lecture-hall priority","(1 = first choice)"],...LECTURE_HALLS.map((h)=>[`Room ${h} priority`, (roomPri&&roomPri[h])||1])];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cRows), "Courses");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(iRows), "Instructors");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sRows), "Setup");
  return wb;
}
function downloadCourseData(courses, teachers, termStart, roomPri) {
  XLSX.writeFile(buildCourseWorkbook(courses, teachers, termStart, roomPri), "Mandakh_Course_Data.xlsx");
}
function findCol(headerRow, ...keys) {
  for (let i=0;i<headerRow.length;i++){ const h=String(headerRow[i]||"").toLowerCase();
    if (keys.some((k)=>h.includes(k))) return i; } return -1;
}
function sheetRows(wb, nameKey) {
  const name = wb.SheetNames.find((n)=>n.toLowerCase().includes(nameKey));
  if (!name) return [];
  return XLSX.utils.sheet_to_json(wb.Sheets[name], { header:1, blankrows:false, defval:"" });
}
function parseCourseWorkbook(wb) {
  const nonEmpty = (r)=> r.filter((x)=>String(x||"").trim()!=="").length;
  const iAll = sheetRows(wb, "instructor");
  const iHdr = iAll.findIndex((r)=> nonEmpty(r)>=3 && findCol(r,"instructor","teacher","багш")>=0 && (findCol(r,"avoid","day off","амрал")>=0));
  const teachers = []; const nameToId = {};
  const dayId = { monday:"mon",tuesday:"tue",wednesday:"wed",thursday:"thu",friday:"fri" };
  if (iHdr>=0) { const H=iAll[iHdr];
    const cN=findCol(H,"instructor","teacher","багш"), cOff=findCol(H,"day off","амрал"),
      cP=[1,2,3,4].map((n)=>findCol(H,`avoid p${n}`,`p${n}`));
    for (let r=iHdr+1;r<iAll.length;r++){ const row=iAll[r]; const nm=String(row[cN]||"").trim(); if (!nm) continue;
      const id = `t${teachers.length+1}`; nameToId[nm]=id;
      const avoid=[]; cP.forEach((ci,k)=>{ if (ci>=0 && yes(row[ci])) avoid.push({ day:"any", period:k+1 }); });
      const offRaw = cOff>=0 ? String(row[cOff]||"").trim().toLowerCase() : "";
      teachers.push({ id, name:nm, off: dayId[offRaw]||null, pref:null, avoid });
    }
  }
  const cAll = sheetRows(wb, "course");
  const cHdr = cAll.findIndex((r)=> nonEmpty(r)>=5 && findCol(r,"course name","name","нэр")>=0 && findCol(r,"length","үргэл")>=0 && findCol(r,"lecture","лекц")>=0);
  if (cHdr<0) throw new Error("Could not find a Courses sheet with the expected columns.");
  const H=cAll[cHdr];
  const col = { name:findCol(H,"course name","name","нэр"), ins:findCol(H,"instructor","teacher","багш"),
    coh:findCol(H,"cohort","group","бүлэг"), len:findCol(H,"length","үргэл"),
    lec:findCol(H,"lecture","лекц"), sem:findCol(H,"seminar","семинар"), bonus:findCol(H,"bonus","нэмэлт"),
    lab:findCol(H,"lab","лаб"), prac:findCol(H,"practical","дадлага"),
    sl:findCol(H,"shared lecture","хамтарсан лекц"), ss:findCol(H,"shared seminar","хамтарсан семинар"), on:findCol(H,"online","онлайн"),
    par:findCol(H,"parallel","зэрэгцээ"), t2:findCol(H,"teacher 2","2-р багш","instructor 2") };
  const validCoh = new Set(COHORTS.map((c)=>c.id));
  const courses = []; let idx=0;
  for (let r=cHdr+1;r<cAll.length;r++){ const row=cAll[r]; const nm=String(row[col.name]||"").trim(); if (!nm) continue;
    idx++;
    const insName=String(row[col.ins]||"").trim();
    let tId = nameToId[insName];
    if (!tId && insName) { tId=`t${teachers.length+1}`; nameToId[insName]=tId; teachers.push({ id:tId, name:insName, off:null, pref:null, avoid:[] }); }
    if (!tId) tId = teachers[0]?.id || "t1";
    const cohorts = String(row[col.coh]||"").split(/[,+/ ]+/).map((x)=>x.trim().toUpperCase()).filter((x)=>validCoh.has(x));
    const num=(ci)=> ci>=0 ? (parseFloat(row[ci])||0) : 0;
    const parallel = col.par>=0 && yes(row[col.par]);
    let t2Id = null;
    if (parallel && col.t2>=0) { const n2=String(row[col.t2]||"").trim(); if (n2) { t2Id = nameToId[n2]; if (!t2Id) { t2Id=`t${teachers.length+1}`; nameToId[n2]=t2Id; teachers.push({ id:t2Id, name:n2, off:null, pref:null, avoid:[] }); } } }
    courses.push({ idx, name:nm, t:tId, cohorts: cohorts.length?cohorts:["1A"], length: col.len>=0?labelToLen(row[col.len]):"full",
      lec:num(col.lec), sem:num(col.sem), bonus:num(col.bonus), lab:num(col.lab), prac:num(col.prac),
      online: col.on>=0 && yes(row[col.on]), combineLec: col.sl>=0 && yes(row[col.sl]), combineSem: col.ss>=0 && yes(row[col.ss]),
      parallel, t2:t2Id });
  }
  if (!courses.length) throw new Error("No course rows found in the Courses sheet.");
  if (!teachers.length) teachers.push({ id:"t1", name:"Instructor 1", off:null, pref:null, avoid:[] });
  let termStart = DEFAULT_TERM_START; const roomPri = { "104":1,"306":2 };
  const sAll = sheetRows(wb, "setup");
  for (const row of sAll) { const key=String(row[0]||"").toLowerCase(); const val=row[1];
    if (key.includes("start date") && val) { const d=(val instanceof Date)?val.toISOString().slice(0,10):String(val).trim(); if (d) termStart=d; }
    LECTURE_HALLS.forEach((h)=>{ if (key.includes(h) && key.includes("priorit")) { const n=parseInt(val); if (n>=1) roomPri[h]=n; } });
  }
  return { courses, teachers, termStart, roomPri };
}

// ============================================================================
//  UI atoms
// ============================================================================
const INK="#1a2438", PAPER="#f7f6f2", LINE="#e3e0d6", THICK="#b0a99c";

function LockBtn({ locked, lang, onClick }) {
  return (
    <button onClick={(e)=>{e.stopPropagation();onClick();}} className="shrink-0 opacity-45 hover:opacity-100 transition-opacity"
      title={locked ? tr(lang,"unlockClass") : tr(lang,"lockClass")} style={locked?{opacity:1,color:"#b45309"}:{}}>
      {locked ? <Lock size={11} strokeWidth={2.6}/> : <Unlock size={11}/>}
    </button>
  );
}

// Full (weekly) card
function ScheduleCard({ s, view, colorBy, lang, courses, teachers, onClick, onToggleLock, onDragStart, half }) {
  const color = cardColor(s, colorBy);
  const online = s.room === "ONLINE";
  const biweekly = s.parity !== "weekly";
  return (
    <div draggable onDragStart={onDragStart} onClick={onClick}
      className="group relative rounded-md px-2 py-1.5 text-[11px] leading-tight cursor-pointer transition-shadow hover:shadow-md flex-1"
      style={{ background: biweekly ? `repeating-linear-gradient(45deg, ${color}0a, ${color}0a 5px, ${color}24 5px, ${color}24 10px)` : `${color}12`, borderLeft:`3px ${biweekly?"dashed":"solid"} ${color}`, color:INK }}
      title={`${courseName(courses,s.courseIdx)} · ${compLabel(lang,s.type)}${biweekly?` · ${s.parity==="odd"?tr(lang,"oddWk"):tr(lang,"evenWk")}`:""}`}>
      <div className="flex items-start justify-between gap-1">
        <span className="font-semibold truncate">{courseName(courses, s.courseIdx)}{s.parallel && <span className="ml-1 text-[9px] font-bold" style={{ color }} title={tr(lang,"parallelNote")}>∥</span>}</span>
        <LockBtn locked={lockedIn(s, half)} lang={lang} onClick={onToggleLock}/>
      </div>
      {s.parallel ? (
        <div className="mt-0.5 text-[10px] opacity-80 space-y-0.5">
          <div className="flex items-center gap-1 flex-wrap">
            <span className="rounded px-1 font-semibold" style={{ background:`${color}22`, color }}>{compTag(lang,s.type)}</span>
            {view!=="cohort" && <span>{s.cohorts.join("+")}</span>}
          </div>
          <div>① {teacherName(teachers,s.ins)} · {online?tr(lang,"online"):s.room}</div>
          <div>② {teacherName(teachers,s.ins2)} · {online?tr(lang,"online"):s.room2}</div>
        </div>
      ) : (
      <div className="mt-0.5 flex items-center gap-1 flex-wrap text-[10px] opacity-80">
        <span className="rounded px-1 font-semibold" style={{ background:`${color}22`, color }}>{compTag(lang,s.type)}</span>
        {s.parity!=="weekly" && <span className="rounded px-1 font-bold uppercase tracking-wide" style={{ background:`${color}2e`, color, fontSize:"8px" }}>⟳ {s.parity==="odd"?tr(lang,"oddWk"):tr(lang,"evenWk")}</span>}
        {view!=="cohort" && <span>{s.cohorts.join("+")}</span>}
        {view!=="instructor" && <span>{teacherName(teachers,s.ins)}</span>}
        {view!=="room" && <span>{online ? tr(lang,"online") : s.room}</span>}
      </div>
      )}
    </div>
  );
}

// Compact card used inside a triangle
function TriCard({ s, align, view, colorBy, lang, courses, teachers, onClick, onToggleLock, onDragStart, half }) {
  const color = cardColor(s, colorBy);
  const online = s.room === "ONLINE";
  const right = align === "right";
  return (
    <div draggable onDragStart={onDragStart} onClick={onClick}
      className="cursor-pointer text-[10px] leading-tight hover:opacity-80"
      style={{ textAlign: right ? "right" : "left" }}
      title={`${courseName(courses,s.courseIdx)} · ${compLabel(lang,s.type)} · ${s.parity==="odd"?tr(lang,"oddWk"):tr(lang,"evenWk")}`}>
      <div className={`flex items-center gap-1 ${right?"flex-row-reverse":""}`}>
        <span className="font-bold truncate" style={{ color:INK, maxWidth:96 }}>{courseName(courses, s.courseIdx)}</span>
        <LockBtn locked={lockedIn(s, half)} lang={lang} onClick={onToggleLock}/>
      </div>
      <div className={`flex items-center gap-1 flex-wrap ${right?"justify-end":""}`} style={{ color:INK, opacity:0.85 }}>
        <span className="rounded px-1 font-semibold" style={{ background:`${color}26`, color }}>{compTag(lang,s.type)}</span>
        {view!=="instructor" && <span>{teacherName(teachers,s.ins)}</span>}
        {view!=="room" && <span>{online ? tr(lang,"online") : s.room}</span>}
      </div>
    </div>
  );
}

// Cell content: weekly cards, or two triangle blocks for biweekly
function CellContent({ list, view, colorBy, lang, courses, teachers, onCard, onLock, setDragId, half }) {
  const weekly = list.filter((s)=>s.parity==="weekly");
  const odd = list.filter((s)=>s.parity==="odd");
  const even = list.filter((s)=>s.parity==="even");
  const full = (s) => (
    <ScheduleCard key={s.id} s={s} view={view} colorBy={colorBy} lang={lang} courses={courses} teachers={teachers} half={half}
      onDragStart={()=>setDragId(s.id)} onClick={()=>onCard(s)} onToggleLock={()=>onLock(s.id)} />
  );
  const tri = (s, align) => (
    <TriCard key={s.id} s={s} align={align} view={view} colorBy={colorBy} lang={lang} courses={courses} teachers={teachers} half={half}
      onDragStart={()=>setDragId(s.id)} onClick={()=>onCard(s)} onToggleLock={()=>onLock(s.id)} />
  );
  if (!odd.length && !even.length) return <div className="flex flex-col gap-1 h-full">{weekly.map(full)}</div>;
  // Only split the cell diagonally when BOTH weeks carry a (different) class. A lone biweekly fills the cell.
  if (!odd.length || !even.length) return <div className="flex flex-col gap-1 h-full">{weekly.map(full)}{[...odd, ...even].map(full)}</div>;
  const oddColor = odd[0] ? cardColor(odd[0], colorBy) : LINE;
  const evenColor = even[0] ? cardColor(even[0], colorBy) : LINE;
  return (
    <div className="flex flex-col gap-1 h-full">
      {weekly.map(full)}
      <div className="relative rounded-md overflow-hidden flex-1" style={{ minHeight:80, border:`1px solid ${LINE}` }}>
        {odd.length>0 && <div className="absolute inset-0" style={{ clipPath:"polygon(0 0,100% 0,0 100%)", background:`${oddColor}14` }}/>}
        {even.length>0 && <div className="absolute inset-0" style={{ clipPath:"polygon(100% 0,100% 100%,0 100%)", background:`${evenColor}14` }}/>}
        <svg className="absolute inset-0 w-full h-full pointer-events-none" preserveAspectRatio="none" viewBox="0 0 100 100">
          <line x1="100" y1="0" x2="0" y2="100" stroke="#c9c3b4" strokeWidth="1.25" vectorEffect="non-scaling-stroke"/>
        </svg>
        {/* odd = top-left */}
        <div className="absolute top-1 left-1.5 right-7 flex flex-col gap-0.5">
          <span className="self-start text-[8px] font-bold px-1 rounded-sm" style={{ background:"#ffffffcc", color:INK }}>{tr(lang,"oddWk")}</span>
          {odd.map((s)=>tri(s,"left"))}
        </div>
        {/* even = bottom-right */}
        <div className="absolute bottom-1 right-1.5 left-7 flex flex-col gap-0.5 items-stretch">
          {even.map((s)=>tri(s,"right"))}
          <span className="self-end text-[8px] font-bold px-1 rounded-sm" style={{ background:"#ffffffcc", color:INK }}>{tr(lang,"evenWk")}</span>
        </div>
      </div>
    </div>
  );
}

function Bar({ label, value }) {
  const col = value>=90 ? "#0d9488" : value>=75 ? "#b45309" : "#be123c";
  return (
    <div>
      <div className="flex justify-between text-[11px] mb-0.5" style={{ color:INK }}><span>{label}</span><span className="font-semibold">{value}%</span></div>
      <div className="h-1.5 rounded-full" style={{ background:LINE }}><div className="h-1.5 rounded-full transition-all" style={{ width:`${value}%`, background:col }}/></div>
    </div>
  );
}
function Panel({ title, icon:Icon, children, accent=INK, onClose }) {
  return (
    <div className="rounded-xl border" style={{ borderColor:LINE, background:"#fff" }}>
      <div className="flex items-center justify-between px-3 py-2" style={{ borderBottom:`1px solid ${LINE}` }}>
        <div className="flex items-center gap-1.5 text-[12px] font-semibold" style={{ color:accent }}><Icon size={14}/> {title}</div>
        {onClose && <button onClick={onClose} className="opacity-40 hover:opacity-100"><X size={14}/></button>}
      </div>
      <div className="p-3">{children}</div>
    </div>
  );
}
function SelectedDetail({ s, lang, courses, teachers, placed, half, onLock, onSetRoom }) {
  const row = (k,v)=>(<div className="flex justify-between text-[11px] py-0.5 gap-3"><span className="opacity-50 whitespace-nowrap">{k}</span><span className="font-medium text-right">{v}</span></div>);
  const lenLabel = { full:tr(lang,"full"), h1:tr(lang,"firstHalf"), h2:tr(lang,"secondHalf") }[s.phase];
  const freqLabel = s.freq==="biweekly" ? `${tr(lang,"biweekly")} (${s.parity==="odd"?tr(lang,"oddWk"):tr(lang,"evenWk")})` : tr(lang,"weekly");
  const rooms = (placed && onSetRoom) ? availableRoomsFor(s, placed, half) : null;
  return (
    <div>
      <div className="font-semibold text-[13px] mb-1">{courseName(courses, s.courseIdx)} · {compLabel(lang,s.type)}</div>
      {row(tr(lang,"length"), lenLabel)}
      {row(tr(lang,"frequency"), freqLabel)}
      {row(tr(lang,"cohorts"), s.cohorts.join(" + "))}
      {s.parallel ? (<>
        {row(`${tr(lang,"section")} ①`, `${teacherName(teachers,s.ins)} · ${s.room==="ONLINE"?tr(lang,"online"):s.room}`)}
        {row(`${tr(lang,"section")} ②`, `${teacherName(teachers,s.ins2)} · ${s.room2==="ONLINE"?tr(lang,"online"):s.room2}`)}
      </>) : (<>
        {row(tr(lang,"instructor"), teacherName(teachers, s.ins))}
        {(rooms && s.roomType!=="online") ? (
          <div className="flex justify-between items-center text-[11px] py-0.5 gap-3">
            <span className="opacity-50 whitespace-nowrap">{tr(lang,"room")}</span>
            <select value={s.room} onChange={(e)=>onSetRoom(s.id, e.target.value)} className="rounded border px-1.5 py-0.5 text-[11px] font-medium text-right" style={{ borderColor:LINE }}>
              {rooms.map((rid)=><option key={rid} value={rid}>{rid}</option>)}
            </select>
          </div>
        ) : row(tr(lang,"room"), s.room==="ONLINE" ? tr(lang,"online") : s.room)}
      </>)}
      {row(tr(lang,"dayPeriod"), `${dayLabel(lang, DAYS.find((x)=>x.id===s.day))} · P${s.period}`)}
      {row(tr(lang,"students"), s.students)}
      <button onClick={onLock} className="mt-2.5 w-full flex items-center justify-center gap-1.5 rounded-md py-1.5 text-[12px] font-semibold"
        style={lockedIn(s, half) ? { background:"#b4530915", color:"#b45309" } : { background:INK, color:PAPER }}>
        {lockedIn(s, half) ? <><Unlock size={13}/> {tr(lang,"unlockClass")}</> : <><Lock size={13}/> {tr(lang,"lockClass")}</>}
      </button>
    </div>
  );
}

// small day+period picker for adding an instructor time-off request
function TeacherAvoidAdd({ lang, onAdd }) {
  const [day, setDay] = useState("any");
  const [period, setPeriod] = useState(1);
  return (
    <span className="inline-flex items-center gap-1">
      <select value={day} onChange={(e)=>setDay(e.target.value)} className="rounded border px-1 py-0.5 text-[10px]" style={{ borderColor:LINE }}>
        <option value="any">{tr(lang,"anyDay")}</option>
        {DAYS.map((d)=><option key={d.id} value={d.id}>{dayLabel(lang,d)}</option>)}
      </select>
      <select value={period} onChange={(e)=>setPeriod(Number(e.target.value))} className="rounded border px-1 py-0.5 text-[10px]" style={{ borderColor:LINE }}>
        {PERIODS.map((p)=><option key={p.id} value={p.id}>P{p.id}</option>)}
      </select>
      <button onClick={()=>onAdd(day, period)} className="rounded-full px-1.5 py-0.5 text-[10px] font-semibold" style={{ background:INK, color:PAPER }}><Plus size={10}/></button>
    </span>
  );
}
// ============================================================================
//  Course data editor
// ============================================================================
function CourseEditor({ lang, courses, setCourses, teachers, setTeachers, termStart, setTermStart, roomPri, setRoomPri, onApply }) {
  const t = (k)=>tr(lang,k);
  const upd = (idx, field, val) => setCourses((prev)=>prev.map((c)=> c.idx===idx ? { ...c, [field]:val } : c));
  const updTeacher = (id, field, val) => setTeachers((prev)=>prev.map((x)=> x.id===id ? { ...x, [field]:val } : x));
  const addAvoid = (id, day, period) => setTeachers((prev)=>prev.map((x)=> x.id===id ? { ...x, avoid:[...(x.avoid||[]), { day, period:Number(period) }] } : x));
  const removeAvoid = (id, i) => setTeachers((prev)=>prev.map((x)=> x.id===id ? { ...x, avoid:(x.avoid||[]).filter((_,j)=>j!==i) } : x));
  const del = (idx) => setCourses((prev)=>prev.filter((c)=>c.idx!==idx));
  const toggleCohort = (idx, co) => setCourses((prev)=>prev.map((c)=>{
    if (c.idx!==idx) return c;
    const has = c.cohorts.includes(co);
    let next = has ? c.cohorts.filter((x)=>x!==co) : [...c.cohorts, co];
    if (next.length===0) next = [co];
    next = COHORTS.filter((x)=>next.includes(x.id)).map((x)=>x.id);
    return { ...c, cohorts:next, combineLec: next.length>1 ? c.combineLec : false };
  }));
  const add = () => setCourses((prev)=>{ const nx = Math.max(0,...prev.map((c)=>c.idx))+1;
    return [...prev, cc(nx, ["1A"], teachers[0].id, "full", { lec:1, sem:1 })]; });
  const addT = () => setTeachers((prev)=>{ const maxN = prev.reduce((m,x)=>{ const k = parseInt(String(x.id).replace(/\D/g,"")) || 0; return k>m?k:m; }, 0); return [...prev, mkTeacher(maxN+1)]; });
  const derivedText = (c) => deriveComponents(c).map((x)=> `${compTag(lang,x.type)}·${x.freq==="biweekly"?tr(lang,"biweekly"):tr(lang,"weekly")}`).join(", ") || "—";
  const numCell = (c, field) => (
    <input type="number" min="0" step="0.5" value={c[field]} onChange={(e)=>upd(c.idx, field, parseFloat(e.target.value)||0)}
      className="w-12 rounded border text-center text-[12px] py-0.5" style={{ borderColor:LINE }} />
  );
  const fileRef = useRef(null);
  const delTeacher = (id) => {
    const used = courses.filter((c)=>c.t===id).length;
    if (used>0) { if (typeof window!=="undefined") window.alert(`${tr(lang,"cantDeleteTeacher")} (${used})`); return; }
    setTeachers((prev)=>prev.filter((x)=>x.id!==id));
  };
  const onExport = () => downloadCourseData(courses, teachers, termStart, roomPri);
  const onImport = (e) => { const f = e.target.files && e.target.files[0]; if (!f) return;
    const rd = new FileReader();
    rd.onload = () => { try { const wb = XLSX.read(rd.result, { type:"array" }); const d = parseCourseWorkbook(wb);
        setCourses(d.courses); setTeachers(d.teachers);
        if (setTermStart && d.termStart) setTermStart(d.termStart);
        if (setRoomPri && d.roomPri) setRoomPri(d.roomPri);
        if (typeof window!=="undefined") window.alert(`${tr(lang,"importOk")}: ${d.courses.length} ${tr(lang,"coursesN")}, ${d.teachers.length} ${tr(lang,"instructorsN")}`);
      } catch (err) { if (typeof window!=="undefined") window.alert(`${tr(lang,"importErr")}\n\n${err.message||err}`); }
      finally { e.target.value=""; } };
    rd.readAsArrayBuffer(f);
  };
  return (
    <div className="rounded-xl border" style={{ borderColor:LINE, background:"#fff" }}>
      <div className="px-4 py-3 flex items-center justify-between flex-wrap gap-3" style={{ borderBottom:`1px solid ${LINE}` }}>
        <div>
          <div className="flex items-center gap-2 font-semibold text-[14px]"><Table2 size={16}/> {t("courseData")}</div>
          <div className="text-[11px] opacity-60 mt-0.5 max-w-2xl">{t("editHint")}</div>
        </div>
        <div className="flex items-center gap-2">
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={onImport} style={{ display:"none" }} />
          <button onClick={()=>fileRef.current&&fileRef.current.click()} className="flex items-center gap-1.5 rounded-md px-3 py-1.5 text-[12px] font-medium" style={{ background:"#fbfaf6", border:`1px solid ${LINE}` }} title={t("importXlsx")}><Upload size={13}/> {t("importXlsx")}</button>
          <button onClick={onExport} className="flex items-center gap-1.5 rounded-md px-3 py-1.5 text-[12px] font-medium" style={{ background:"#fbfaf6", border:`1px solid ${LINE}` }} title={t("exportXlsx")}><Download size={13}/> {t("exportXlsx")}</button>
          <button onClick={()=>{ setCourses(SEED_COURSES); setTeachers(SEED_TEACHERS); }} className="flex items-center gap-1.5 rounded-md px-3 py-1.5 text-[12px] font-medium" style={{ background:"#fbfaf6", border:`1px solid ${LINE}` }} title={t("resetSample")}><RotateCcw size={13}/> {t("resetSample")}</button>
          <button onClick={()=>{ if (typeof window==="undefined" || window.confirm(t("confirmClear"))) { setCourses([]); } }} className="flex items-center gap-1.5 rounded-md px-3 py-1.5 text-[12px] font-medium" style={{ background:"#fff", border:`1px solid ${LINE}`, color:"#be123c" }} title={t("clearAll")}><Trash2 size={13}/> {t("clearAll")}</button>
          <button onClick={add} className="flex items-center gap-1.5 rounded-md px-3 py-1.5 text-[12px] font-medium" style={{ background:"#fbfaf6", border:`1px solid ${LINE}` }}><Plus size={13}/> {t("addCourse")}</button>
          <button onClick={onApply} className="flex items-center gap-1.5 rounded-md px-3 py-1.5 text-[12px] font-semibold" style={{ background:INK, color:PAPER }}><RotateCw size={13}/> {t("applyRegen")}</button>
        </div>
      </div>
      <div className="px-4 py-3 flex items-center gap-3 flex-wrap" style={{ borderBottom:`1px solid ${LINE}`, background:"#fbfaf6" }}>
        <div className="flex items-center gap-2 font-semibold text-[12px]"><Calendar size={14}/> {t("semesterStart")}</div>
        <input type="date" value={termStart||""} onChange={(e)=>setTermStart(e.target.value)} className="rounded border px-2 py-1 text-[12px]" style={{ borderColor:LINE }} />
        <span className="text-[12px] rounded-full px-2.5 py-0.5 font-medium" style={{ background:"#eef2ff", color:INK }}>{t("weekStartsOn")}: {dayLabel(lang, WEEKDAYS[weekdayIndex(termStart)])}</span>
        <span className="text-[11px] opacity-60 flex-1 min-w-[180px]">{t("startDateNote")}</span>
      </div>
      <div className="px-4 py-3 flex items-center gap-3 flex-wrap" style={{ borderBottom:`1px solid ${LINE}`, background:"#fbfaf6" }}>
        <div className="flex items-center gap-2 font-semibold text-[12px]"><DoorOpen size={14}/> {t("roomPriorityTitle")}</div>
        {LECTURE_HALLS.map((h)=>(
          <label key={h} className="flex items-center gap-1.5 text-[12px] rounded-md px-2 py-1" style={{ background:"#fff", border:`1px solid ${LINE}` }}>
            <span className="font-semibold">{h}</span>
            <input type="number" min="1" max="9" value={(roomPri&&roomPri[h])||1}
              onChange={(e)=>setRoomPri((prev)=>({ ...prev, [h]:Math.max(1, parseInt(e.target.value)||1) }))}
              className="w-12 rounded border text-center text-[12px] py-0.5" style={{ borderColor:LINE }} />
          </label>
        ))}
        <span className="text-[11px] opacity-60 flex-1 min-w-[180px]">{t("roomPriorityNote")}</span>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-left border-collapse text-[12px]" style={{ minWidth:1040 }}>
          <thead>
            <tr style={{ background:"#fbfaf6" }}>
              {[t("colName"), t("colTeacher"), t("colCohorts"), t("colLength")].map((h)=>(<th key={h} className="px-2.5 py-2 font-semibold whitespace-nowrap" style={{ borderBottom:`1px solid ${LINE}` }}>{h}</th>))}
              <th className="px-2.5 py-2 font-semibold text-center" colSpan={5} style={{ borderBottom:`1px solid ${LINE}`, borderLeft:`1px solid ${LINE}` }}>{t("colCredits")}</th>
              <th className="px-2.5 py-2 font-semibold" style={{ borderBottom:`1px solid ${LINE}`, borderLeft:`1px solid ${LINE}` }}>{t("colGenerates")}</th>
              <th style={{ borderBottom:`1px solid ${LINE}` }}></th>
            </tr>
            <tr style={{ background:"#fbfaf6" }}>
              <th colSpan={4} style={{ borderBottom:`1px solid ${LINE}` }}></th>
              {["cLecture","cSeminar","cBonus","cLab","cPractical"].map((k,i)=>(<th key={k} className="px-1.5 py-1 font-medium text-center text-[10px] opacity-70" style={{ borderBottom:`1px solid ${LINE}`, borderLeft:i===0?`1px solid ${LINE}`:"none" }}>{t(k)}</th>))}
              <th style={{ borderBottom:`1px solid ${LINE}`, borderLeft:`1px solid ${LINE}` }}></th>
              <th style={{ borderBottom:`1px solid ${LINE}` }}></th>
            </tr>
          </thead>
          <tbody>
            {courses.map((c)=>(
              <tr key={c.idx} className="align-middle">
                <td className="px-2.5 py-1.5" style={{ borderBottom:`1px solid ${LINE}` }}><input value={c.name} onChange={(e)=>upd(c.idx,"name",e.target.value)} className="w-36 rounded border px-1.5 py-0.5 text-[12px]" style={{ borderColor:LINE }} /></td>
                <td className="px-2.5 py-1.5" style={{ borderBottom:`1px solid ${LINE}` }}>
                  <select value={c.t} onChange={(e)=>upd(c.idx,"t",e.target.value)} className="rounded border px-1 py-0.5 text-[12px]" style={{ borderColor:LINE }}>
                    {teachers.map((tt)=><option key={tt.id} value={tt.id}>{tt.name}</option>)}
                  </select>
                </td>
                <td className="px-2.5 py-1.5" style={{ borderBottom:`1px solid ${LINE}`, minWidth:150 }}>
                  <div className="flex gap-0.5 flex-wrap">
                    {COHORTS.map((co)=>{ const on = c.cohorts.includes(co.id);
                      return (<button key={co.id} onClick={()=>toggleCohort(c.idx, co.id)} className="rounded px-1 text-[10px] font-semibold transition-colors"
                        style={on ? { background:`${COHORT_COLORS[co.id]}22`, color:COHORT_COLORS[co.id] } : { background:"#f0eee7", color:"#b8b2a4" }}>{co.id}</button>);
                    })}
                  </div>
                  {c.cohorts.length>1 && (
                    <div className="mt-1 space-y-0.5">
                      <label className="flex items-center gap-1 text-[10px] opacity-75 cursor-pointer">
                        <input type="checkbox" checked={c.combineLec} onChange={(e)=>upd(c.idx,"combineLec",e.target.checked)} /> {t("sharedLecture")}
                      </label>
                      <label className="flex items-center gap-1 text-[10px] opacity-75 cursor-pointer">
                        <input type="checkbox" checked={c.combineSem} onChange={(e)=>upd(c.idx,"combineSem",e.target.checked)} /> {t("sharedSeminar")}
                      </label>
                    </div>
                  )}
                  {c.sem>0 && (
                    <div className="mt-1">
                      <label className="flex items-center gap-1 text-[10px] cursor-pointer" style={{ color:c.parallel?"#7c3aed":undefined, opacity:c.parallel?1:0.75 }} title={t("parallelHint")}>
                        <input type="checkbox" checked={!!c.parallel} onChange={(e)=>upd(c.idx,"parallel",e.target.checked)} /> {t("parallelFlag")}
                      </label>
                      {c.parallel && (
                        <div className="mt-0.5 flex items-center gap-1 text-[10px]">
                          <span className="opacity-50">②</span>
                          <select value={c.t2||""} onChange={(e)=>upd(c.idx,"t2",e.target.value||null)} className="rounded border px-1 py-0.5 text-[10px]" style={{ borderColor:LINE }}>
                            <option value="">{t("teacher2")}</option>
                            {teachers.map((tt)=><option key={tt.id} value={tt.id}>{tt.name}</option>)}
                          </select>
                        </div>
                      )}
                    </div>
                  )}
                </td>
                <td className="px-2.5 py-1.5" style={{ borderBottom:`1px solid ${LINE}` }}>
                  <select value={c.length} onChange={(e)=>upd(c.idx,"length",e.target.value)} className="rounded border px-1 py-0.5 text-[12px]" style={{ borderColor:LINE }}>
                    <option value="full">{t("full")}</option><option value="h1">{t("firstHalf")}</option><option value="h2">{t("secondHalf")}</option>
                  </select>
                  <label className="flex items-center gap-1 mt-1 text-[10px] opacity-75 cursor-pointer" title={t("onlineHint")}>
                    <input type="checkbox" checked={!!c.online} onChange={(e)=>upd(c.idx,"online",e.target.checked)} /> {t("onlineFlag")}
                  </label>
                </td>
                <td className="px-1.5 py-1.5 text-center" style={{ borderBottom:`1px solid ${LINE}`, borderLeft:`1px solid ${LINE}` }}>{numCell(c,"lec")}</td>
                <td className="px-1.5 py-1.5 text-center" style={{ borderBottom:`1px solid ${LINE}` }}>{numCell(c,"sem")}</td>
                <td className="px-1.5 py-1.5 text-center" style={{ borderBottom:`1px solid ${LINE}` }}>{numCell(c,"bonus")}</td>
                <td className="px-1.5 py-1.5 text-center" style={{ borderBottom:`1px solid ${LINE}` }}>{numCell(c,"lab")}</td>
                <td className="px-1.5 py-1.5 text-center" style={{ borderBottom:`1px solid ${LINE}` }}>{numCell(c,"prac")}</td>
                <td className="px-2.5 py-1.5 text-[10px] opacity-70 whitespace-nowrap" style={{ borderBottom:`1px solid ${LINE}`, borderLeft:`1px solid ${LINE}` }}>{derivedText(c)}</td>
                <td className="px-2 py-1.5 text-center" style={{ borderBottom:`1px solid ${LINE}` }}><button onClick={()=>del(c.idx)} className="opacity-40 hover:opacity-100" style={{ color:"#be123c" }}><Trash2 size={13}/></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="px-4 py-2 text-[11px] opacity-60 space-y-1" style={{ borderTop:`1px solid ${LINE}` }}>
        <div>{t("cohortNote")}</div>
        <div>{t("creditRuleNote")}</div>
      </div>
      <div className="px-4 py-3" style={{ borderTop:`1px solid ${LINE}` }}>
        <div className="flex items-center justify-between mb-2">
          <div className="flex items-center gap-2 font-semibold text-[13px]"><GraduationCap size={15}/> {t("teachersTitle")}</div>
          <button onClick={addT} className="flex items-center gap-1.5 rounded-md px-2.5 py-1 text-[12px] font-medium" style={{ background:"#fbfaf6", border:`1px solid ${LINE}` }}><Plus size={12}/> {t("addTeacher")}</button>
        </div>
        <div className="text-[11px] opacity-60 mb-2">{t("requestNote")}</div>
        <div className="grid gap-2" style={{ gridTemplateColumns:"repeat(auto-fill, minmax(280px, 1fr))" }}>
          {teachers.map((tt)=>(
            <div key={tt.id} className="rounded-md px-2 py-1.5 space-y-1.5" style={{ background:"#fbfaf6" }}>
              <div className="flex items-center gap-2">
                <input value={tt.name} onChange={(e)=>updTeacher(tt.id,"name",e.target.value)} className="flex-1 min-w-0 rounded border px-1.5 py-0.5 text-[12px]" style={{ borderColor:LINE }} />
                <span className="text-[10px] opacity-50 whitespace-nowrap">{t("dayOff")}</span>
                <select value={tt.off||""} onChange={(e)=>updTeacher(tt.id,"off",e.target.value||null)} className="rounded border px-1 py-0.5 text-[11px]" style={{ borderColor:LINE }}>
                  <option value="">{t("none")}</option>{DAYS.map((d)=><option key={d.id} value={d.id}>{dayLabel(lang,d)}</option>)}
                </select>
                <button onClick={()=>delTeacher(tt.id)} title={t("deleteTeacher")} className="opacity-40 hover:opacity-100" style={{ color:"#be123c" }}><Trash2 size={13}/></button>
              </div>
              <div className="flex items-center gap-1 flex-wrap">
                <span className="text-[10px] opacity-50">{t("softRequests")}:</span>
                {(tt.avoid||[]).length===0 && <span className="text-[10px] opacity-40">—</span>}
                {(tt.avoid||[]).map((a,i)=>(
                  <button key={i} onClick={()=>removeAvoid(tt.id,i)} className="flex items-center gap-1 rounded-full px-1.5 py-0.5 text-[10px]" style={{ background:"#fee2e2", color:"#b91c1c" }} title={t("softRequests")}>
                    {a.day==="any" ? tr(lang,"anyDay") : dayLabel(lang, WEEKDAYS.find((d)=>d.id===a.day)||{en:a.day,mn:a.day})} · P{a.period} <X size={10}/>
                  </button>
                ))}
                <TeacherAvoidAdd lang={lang} onAdd={(day,period)=>addAvoid(tt.id, day, period)} />
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ============================================================================
//  App
// ============================================================================
function AuthModal({ lang, msg, onClose, onAuth }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  return (
    <div className="no-print fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background:"#1a2438aa" }} onClick={onClose}>
      <div className="rounded-xl max-w-sm w-full" style={{ background:PAPER }} onClick={(e)=>e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-3" style={{ background:INK, color:PAPER }}>
          <div className="flex items-center gap-2 font-semibold text-[14px]"><User size={15}/> {tr(lang,"signIn")}</div>
          <button onClick={onClose} className="rounded-md p-1" style={{ background:"#ffffff1a" }}><X size={16}/></button>
        </div>
        <div className="p-5 space-y-3">
          <div className="text-[11px] opacity-60">{tr(lang,"authHint")}</div>
          <div>
            <label className="text-[11px] opacity-60">{tr(lang,"emailLbl")}</label>
            <input type="email" value={email} onChange={(e)=>setEmail(e.target.value)} className="w-full rounded border px-2 py-1.5 text-[13px] mt-0.5" style={{ borderColor:LINE }} autoFocus />
          </div>
          <div>
            <label className="text-[11px] opacity-60">{tr(lang,"passwordLbl")}</label>
            <input type="password" value={password} onChange={(e)=>setPassword(e.target.value)} className="w-full rounded border px-2 py-1.5 text-[13px] mt-0.5" style={{ borderColor:LINE }} />
          </div>
          {msg && <div className="text-[11px]" style={{ color: msg==="…"?"#6b7280":"#b91c1c" }}>{msg}</div>}
          <div className="flex gap-2 pt-1">
            <button onClick={()=>onAuth("in", email, password)} className="flex-1 rounded-md py-2 text-[13px] font-semibold" style={{ background:INK, color:PAPER }}>{tr(lang,"signIn")}</button>
            <button onClick={()=>onAuth("up", email, password)} className="flex-1 rounded-md py-2 text-[13px] font-semibold" style={{ background:"#fbfaf6", border:`1px solid ${LINE}` }}>{tr(lang,"signUp")}</button>
          </div>
        </div>
      </div>
    </div>
  );
}
function AdminModal({ lang, list, users, selfId, onClose, onLoad, onRefresh, onRefreshUsers, onSetAdmin }) {
  const [q, setQ] = useState("");
  const [tab, setTab] = useState("schedules");
  const badge = { draft:"#6b7280", checkpoint:"#16a34a", final:"#7c3aed", autosave:"#f59e0b" };
  const rows = (list||[]).filter((r)=>{ const s=(q||"").toLowerCase(); return !s || (r.owner_email||"").toLowerCase().includes(s) || (r.name||"").toLowerCase().includes(s); });
  const urows = (users||[]).filter((r)=>{ const s=(q||"").toLowerCase(); return !s || (r.email||"").toLowerCase().includes(s); });
  const TabBtn = ({ id, label }) => (
    <button onClick={()=>{ setTab(id); setQ(""); if (id==="users") onRefreshUsers(); }} className="rounded-md px-2.5 py-1 text-[11px] font-medium" style={ tab===id ? { background:"#ffffff33" } : { background:"transparent", opacity:0.7 }}>{label}</button>
  );
  return (
    <div className="no-print fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background:"#1a2438aa" }} onClick={onClose}>
      <div className="rounded-xl max-w-2xl w-full max-h-[85vh] overflow-auto" style={{ background:PAPER }} onClick={(e)=>e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-3 sticky top-0" style={{ background:"#4c1d95", color:PAPER }}>
          <div className="flex items-center gap-3">
            <div className="flex items-center gap-2 font-semibold text-[14px]"><Gauge size={15}/> {tr(lang,"adminTitle")}</div>
            <div className="flex items-center gap-1"><TabBtn id="schedules" label={tr(lang,"adminTabSchedules")}/><TabBtn id="users" label={tr(lang,"adminTabUsers")}/></div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={tab==="users"?onRefreshUsers:onRefresh} className="rounded-md px-2 py-1 text-[11px]" style={{ background:"#ffffff1a" }}>{tr(lang,"refresh")}</button>
            <button onClick={onClose} className="rounded-md p-1" style={{ background:"#ffffff1a" }}><X size={16}/></button>
          </div>
        </div>
        <div className="p-5 space-y-3">
          <p className="text-[11px] opacity-60">{tab==="users" ? tr(lang,"adminUsersDesc") : tr(lang,"adminDesc")}</p>
          <input value={q} onChange={(e)=>setQ(e.target.value)} placeholder={tab==="users"?tr(lang,"adminUserSearch"):tr(lang,"adminSearch")} className="w-full rounded border px-2 py-1.5 text-[13px]" style={{ borderColor:LINE }} />
          {tab==="schedules" ? (
            rows.length===0 ? <div className="text-[12px] opacity-50 text-center py-6">{tr(lang,"adminEmpty")}</div> : (
              <div className="space-y-1.5">
                {rows.map((row)=>(
                  <div key={row.id} className="flex items-center gap-2 rounded-lg px-3 py-2" style={{ background:"#fff", border:`1px solid ${LINE}` }}>
                    <span className="rounded px-1.5 py-0.5 text-[9px] font-bold uppercase" style={{ background:`${badge[row.status]||"#6b7280"}22`, color:badge[row.status]||"#6b7280" }}>{row.status}</span>
                    <div className="flex-1 min-w-0">
                      <div className="text-[13px] font-medium truncate">{row.owner_email || tr(lang,"unknownUser")}</div>
                      <div className="text-[10px] opacity-50 truncate">{row.name} · {new Date(row.updated_at).toLocaleString()}</div>
                    </div>
                    <button onClick={()=>onLoad(row.id)} className="rounded-md px-2.5 py-1 text-[12px] font-semibold shrink-0" style={{ background:INK, color:PAPER }}>{tr(lang,"loadBtn")}</button>
                  </div>
                ))}
              </div>
            )
          ) : (
            urows.length===0 ? <div className="text-[12px] opacity-50 text-center py-6">{tr(lang,"adminEmpty")}</div> : (
              <div className="space-y-1.5">
                {urows.map((u)=>(
                  <div key={u.id} className="flex items-center gap-2 rounded-lg px-3 py-2" style={{ background:"#fff", border:`1px solid ${LINE}` }}>
                    {u.is_admin && <span className="rounded px-1.5 py-0.5 text-[9px] font-bold uppercase" style={{ background:"#7c3aed22", color:"#7c3aed" }}>{tr(lang,"adminBadge")}</span>}
                    <div className="flex-1 min-w-0">
                      <div className="text-[13px] font-medium truncate">{u.email}{u.id===selfId && <span className="opacity-45 font-normal"> · {tr(lang,"you")}</span>}</div>
                      <div className="text-[10px] opacity-50 truncate">{tr(lang,"joined")} {new Date(u.created_at).toLocaleDateString()}</div>
                    </div>
                    {u.id===selfId ? <span className="text-[11px] opacity-40 shrink-0">—</span> :
                      u.is_admin
                        ? <button onClick={()=>onSetAdmin(u.id, false)} className="rounded-md px-2.5 py-1 text-[12px] font-medium shrink-0" style={{ background:"#be123c15", color:"#be123c" }}>{tr(lang,"removeAdmin")}</button>
                        : <button onClick={()=>onSetAdmin(u.id, true)} className="rounded-md px-2.5 py-1 text-[12px] font-semibold shrink-0" style={{ background:"#7c3aed", color:PAPER }}>{tr(lang,"makeAdmin")}</button>}
                  </div>
                ))}
              </div>
            )
          )}
        </div>
      </div>
    </div>
  );
}
function CloudModal({ lang, email, list, termStart, onClose, onSave, onLoad, onDelete, onSignOut }) {
  const [name, setName] = useState("");
  const [status, setStatus] = useState("checkpoint");
  const badge = { draft:"#6b7280", checkpoint:"#16a34a", final:"#7c3aed", autosave:"#f59e0b" };
  return (
    <div className="no-print fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background:"#1a2438aa" }} onClick={onClose}>
      <div className="rounded-xl max-w-xl w-full max-h-[85vh] overflow-auto" style={{ background:PAPER }} onClick={(e)=>e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-3 sticky top-0" style={{ background:INK, color:PAPER }}>
          <div className="flex items-center gap-2 font-semibold text-[14px]"><History size={15}/> {tr(lang,"cloudTitle")}</div>
          <div className="flex items-center gap-2">
            <span className="text-[11px] opacity-70 hidden sm:inline">{email}</span>
            <button onClick={onSignOut} className="rounded-md px-2 py-1 text-[11px]" style={{ background:"#ffffff1a" }}>{tr(lang,"signOut")}</button>
            <button onClick={onClose} className="rounded-md p-1" style={{ background:"#ffffff1a" }}><X size={16}/></button>
          </div>
        </div>
        <div className="p-5 space-y-4">
          <div className="rounded-lg p-3 flex flex-wrap items-end gap-2" style={{ background:"#fbfaf6", border:`1px solid ${LINE}` }}>
            <div className="flex-1 min-w-[160px]">
              <label className="text-[11px] opacity-60">{tr(lang,"scheduleName")}</label>
              <input value={name} onChange={(e)=>setName(e.target.value)} placeholder={`Fall ${termStart||""}`} className="w-full rounded border px-2 py-1.5 text-[13px] mt-0.5" style={{ borderColor:LINE }} />
            </div>
            <select value={status} onChange={(e)=>setStatus(e.target.value)} className="rounded border px-2 py-1.5 text-[12px]" style={{ borderColor:LINE }}>
              <option value="draft">{tr(lang,"statusDraft")}</option>
              <option value="checkpoint">{tr(lang,"statusCheckpoint")}</option>
              <option value="final">{tr(lang,"statusFinal")}</option>
            </select>
            <button onClick={()=>{ onSave(name, status); setName(""); }} className="rounded-md px-3 py-2 text-[13px] font-semibold flex items-center gap-1.5" style={{ background:INK, color:PAPER }}><Save size={13}/> {tr(lang,"saveToCloud")}</button>
          </div>
          {list.length===0 ? (
            <div className="text-[12px] opacity-50 text-center py-6">{tr(lang,"noSaved")}</div>
          ) : (
            <div className="space-y-1.5">
              {list.map((row)=>(
                <div key={row.id} className="flex items-center gap-2 rounded-lg px-3 py-2" style={{ background:"#fff", border:`1px solid ${LINE}` }}>
                  <span className="rounded px-1.5 py-0.5 text-[9px] font-bold uppercase" style={{ background:`${badge[row.status]||"#6b7280"}22`, color:badge[row.status]||"#6b7280" }}>{tr(lang,"status"+(row.status||"draft").charAt(0).toUpperCase()+(row.status||"draft").slice(1))||row.status}</span>
                  <div className="flex-1 min-w-0">
                    <div className="text-[13px] font-medium truncate">{row.name}</div>
                    <div className="text-[10px] opacity-50">{new Date(row.updated_at).toLocaleString()}</div>
                  </div>
                  <button onClick={()=>onLoad(row.id)} className="rounded-md px-2.5 py-1 text-[12px] font-semibold" style={{ background:INK, color:PAPER }}>{tr(lang,"loadBtn")}</button>
                  <button onClick={()=>onDelete(row.id)} className="opacity-40 hover:opacity-100" style={{ color:"#be123c" }}><Trash2 size={14}/></button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
export default function App() {
  const [lang, setLang] = useState("en");
  const [screen, setScreen] = useState("timetable");
  const [semView, setSemView] = useState("h1");
  const [colorBy, setColorBy] = useState("cohort");
  const [termStart, setTermStart] = useState(DEFAULT_TERM_START);
  const [roomPri, setRoomPri] = useState(ROOM_PRI);
  DAYS = weekOrderFromStart(termStart); // active teaching-week order follows the start date
  ROOM_PRI = roomPri;                   // active lecture-hall preference follows the editor
  const [courses, setCourses] = useState(SEED_COURSES);
  const [teachers, setTeachers] = useState(SEED_TEACHERS);
  const [placed, setPlaced] = useState(SEED_PLACED);
  const [unplaced, setUnplaced] = useState([]);
  const [candidates, setCandidates] = useState(()=>[{ placed:SEED_PLACED, unplaced:[], score:scoreSchedule(SEED_PLACED, [], SEED_TEACHERS) }]);
  const [activeCand, setActiveCand] = useState(0);
  const [generating, setGenerating] = useState(false);
  const [genStep, setGenStep] = useState("gen1");
  const [view, setView] = useState("master");
  const [filterId, setFilterId] = useState("1A");
  const [selected, setSelected] = useState(null);
  const [toast, setToast] = useState(null);
  const [versions, setVersions] = useState([]);
  const [dragId, setDragId] = useState(null);
  const [hoverCell, setHoverCell] = useState(null);
  const [showData, setShowData] = useState(false);
  const [rules, setRules] = useState(DEFAULT_RULES);
  const [showRules, setShowRules] = useState(false);
  const [session, setSession] = useState(null);
  const [showAuth, setShowAuth] = useState(false);
  const [showCloud, setShowCloud] = useState(false);
  const [cloudList, setCloudList] = useState([]);
  const [isAdmin, setIsAdmin] = useState(false);
  const [showAdmin, setShowAdmin] = useState(false);
  const [adminList, setAdminList] = useState([]);
  const [adminUsers, setAdminUsers] = useState([]);
  const [showSaveDialog, setShowSaveDialog] = useState(false);
  const [saveName, setSaveName] = useState("");
  const [saveStatus, setSaveStatus] = useState("checkpoint");
  const [authMsg, setAuthMsg] = useState("");
  const t = useCallback((k)=>tr(lang,k), [lang]);

  const score = useMemo(()=> placed.length ? scoreSchedule(placed, unplaced, teachers, semView) : null, [placed, unplaced, teachers, semView]);
  const overloaded = useMemo(()=>{ const load = cohortLoad(courses); const out=[]; const TH = SLOTS_PER_HALF-1;
    for (const c of COHORTS) { const h1=Math.round(load[c.id].h1), h2=Math.round(load[c.id].h2);
      if (h1>=TH || h2>=TH) out.push({ id:c.id, h1, h2 }); }
    return out; }, [courses]);
  const conflicts = useMemo(()=> collisionCells(placed, semView), [placed, semView]);
  const repeats = useMemo(()=>{ const seen={}; const out=[];
    placed.filter((s)=>phaseVisible(s.phase, semView)).forEach((s)=>s.cohorts.forEach((c)=>{ const k=c+"|"+s.day+"|"+s.courseIdx+"|"+s.type; (seen[k]=seen[k]||[]).push(s); }));
    for (const k in seen) { const g=seen[k]; if (new Set(g.map((s)=>s.period)).size>1) { const [c,d,ci]=k.split("|"); out.push({ cohort:c, day:d, courseIdx:+ci, ins:g[0].ins, periods:[...new Set(g.map((s)=>s.period))].sort() }); } }
    return out; }, [placed, semView]);
  const hasHalves = useMemo(()=> placed.some((s)=>s.phase!=="full"), [placed]);

  const runGenerate = useCallback((mode) => { // "all" (keep locks) | "half" (regenerate current half; keep other half + full-sem + this half's locks) | "fresh" (no locks)
    setGenerating(true);
    const locks = (mode==="fresh" ? []
      : mode==="half" ? placed.filter((s)=> !phaseVisible(s.phase, semView) || s.phase==="full" || lockedIn(s, semView))
      : placed.filter((s)=> lockedIn(s,"h1") || lockedIn(s,"h2"))
      ).map((s)=>({ ...s, locked:true })); // mark as pre-placed so the solver + repair leave them fixed
    const steps = ["gen1","gen2","gen3","gen4","gen5"]; let i=0; setGenStep(steps[0]);
    const iv = setInterval(()=>{ i=(i+1)%steps.length; setGenStep(steps[i]); }, 200);
    setTimeout(async ()=>{
      const cands = await generateCandidatesAsync(courses, teachers, locks, 8, rules, undefined, Date.now()); // per-Generate seed base → each click re-rolls
      clearInterval(iv);
      if (cands.length) { setCandidates(cands); setActiveCand(0); setPlaced(clone(cands[0].placed)); setUnplaced(cands[0].unplaced); }
      setGenerating(false); setSelected(null);
    }, 60);
  }, [placed, courses, teachers, rules, semView]);

  // embedded sample schedule shows on load; Generate re-solves on demand

  const showToast = (status, text) => { setToast({ status, text }); setTimeout(()=>setToast(null), 2600); };
  const toggleLock = (id) => { const key = semView==="h1" ? "lockH1" : "lockH2";
    setPlaced((prev)=>prev.map((s)=> s.id===id ? { ...s, [key]:!lockedIn(s, semView) } : s));
    setSelected((sel)=> sel && sel.id===id ? { ...sel, [key]:!lockedIn(sel, semView) } : sel); };
  const setRoomManual = (id, room) => { setPlaced((prev)=>prev.map((s)=> s.id===id ? { ...s, room, locked:true } : s)); setSelected((sel)=> sel && sel.id===id ? { ...sel, room, locked:true } : sel); };
  const halfSessions = useMemo(()=> placed.filter((s)=>phaseVisible(s.phase, semView)), [placed, semView]);
  const halfAllLocked = halfSessions.length>0 && halfSessions.every((s)=>lockedIn(s, semView));
  const lockHalf = () => { const key = semView==="h1" ? "lockH1" : "lockH2"; const target = !halfAllLocked; setPlaced((prev)=>prev.map((s)=> phaseVisible(s.phase, semView) ? { ...s, [key]:target } : s)); showToast("valid", target ? t("lockedHalf") : t("unlockedHalf")); };

  const selectCandidate = (i) => { setActiveCand(i); setPlaced(clone(candidates[i].placed)); setUnplaced(candidates[i].unplaced); setSelected(null); };

  const moveSession = (id, d, p, targetCohort) => {
    const s = placed.find((x)=>x.id===id); if (!s) return;
    if (lockedIn(s, semView)) { showToast("conflict", t("lockedMove")); return; }
    if (targetCohort && !s.cohorts.includes(targetCohort)) { showToast("conflict", t("wrongColumn")); return; } // no horizontal moves into another group's column
    const v = validateMove(s, d, p, placed, lang, teachers, semView);
    if (v.status==="conflict") { showToast("conflict", v.reasons.find((r)=>r.level==="conflict").text); return; }
    setPlaced((prev)=>prev.map((x)=> x.id===id ? { ...x, day:d, period:p, room:v.room, room2:v.room2 } : x));
    showToast(v.status, v.status==="warning" ? v.reasons.find((r)=>r.level==="warning").text : t("moved"));
  };

  const saveVersion = () => {
    const v = { id:versions.length+1, score:score?.overall??0, when:new Date().toLocaleTimeString([], { hour:"2-digit", minute:"2-digit" }),
      snapshot:{ placed:clone(placed), unplaced:clone(unplaced), courses:clone(courses), teachers:clone(teachers), termStart, roomPri:clone(roomPri), rules:clone(rules) } };
    setVersions((prev)=>[...prev, v]); showToast("valid", `${t("savedAs")} V${v.id}`);
  };
  const restoreVersion = (v) => { const s = v.snapshot;
    if (s && s.placed) { setPlaced(clone(s.placed)); setUnplaced(clone(s.unplaced||[]));
      if (s.courses) setCourses(clone(s.courses)); if (s.teachers) setTeachers(clone(s.teachers));
      if (s.termStart) setTermStart(s.termStart); if (s.roomPri) setRoomPri(clone(s.roomPri)); if (s.rules) setRules(clone(s.rules)); }
    else { setPlaced(clone(v.snapshot)); setUnplaced([]); } // back-compat with schedule-only snapshots
    showToast("valid", `${t("restoredV")} V${v.id}`);
  };

  // ---- Supabase auth + cloud schedules ----
  useEffect(()=>{ if (!supabase) return; let sub;
    supabase.auth.getSession().then(({ data })=>setSession(data.session)).catch(()=>{});
    sub = supabase.auth.onAuthStateChange((_e, s)=>setSession(s)).data?.subscription;
    return ()=>{ if (sub) sub.unsubscribe(); };
  }, []);
  const refreshCloud = useCallback(async ()=>{ if (!supabase || !session) return;
    const { data, error } = await supabase.from("schedules").select("id,name,term,status,updated_at").order("updated_at",{ ascending:false });
    if (!error) setCloudList(data||[]);
  }, [session]);
  useEffect(()=>{ if (session) refreshCloud(); else setCloudList([]); }, [session, refreshCloud]);
  // admin detection: a row in the `admins` table (visible to you via its own RLS policy) means you're an admin
  useEffect(()=>{ let alive=true;
    if (!supabase || !session) { setIsAdmin(false); return; }
    supabase.from("admins").select("user_id").eq("user_id", session.user.id).maybeSingle()
      .then(({ data })=>{ if (alive) setIsAdmin(!!data); }).catch(()=>{});
    return ()=>{ alive=false; };
  }, [session]);
  const refreshAdmin = useCallback(async ()=>{ if (!supabase || !isAdmin) return;
    const { data, error } = await supabase.rpc("admin_schedules");
    if (!error) setAdminList(data||[]);
  }, [isAdmin]);
  const refreshUsers = useCallback(async ()=>{ if (!supabase || !isAdmin) return;
    const { data, error } = await supabase.rpc("admin_users");
    if (!error) setAdminUsers(data||[]);
  }, [isAdmin]);
  const setAdminFor = useCallback(async (uuid, make)=>{ if (!supabase || !isAdmin) return;
    const { error } = await supabase.rpc("set_admin", { target:uuid, make });
    if (error) showToast("conflict", error.message); else { showToast("valid", make ? t("madeAdmin") : t("removedAdmin")); refreshUsers(); }
  }, [isAdmin, refreshUsers, lang]);
  const doAuth = async (mode, email, password)=>{ if (!supabase) return; setAuthMsg("…");
    const fn = mode==="up" ? supabase.auth.signUp({ email, password }) : supabase.auth.signInWithPassword({ email, password });
    const { error } = await fn;
    if (error) setAuthMsg(error.message);
    else { setAuthMsg(""); setShowAuth(false); }
  };
  const signOut = async ()=>{ if (supabase) await supabase.auth.signOut(); setSession(null); setShowCloud(false); };
  const cloudSnapshot = ()=>({ placed, unplaced, courses, teachers, termStart, roomPri, rules });
  // ---------- auto-save / unsaved-change protection ----------
  const AUTOSAVE_KEY = "mandakh_scheduler_autosave_v1";
  const snapStr = useCallback(()=>JSON.stringify(cloudSnapshot()), [placed, unplaced, courses, teachers, termStart, roomPri, rules]);
  const savedSnapRef = useRef(null);
  const autoRowRef = useRef(null);
  const [dirty, setDirty] = useState(false);
  const [restore, setRestore] = useState(null);
  useEffect(()=>{ if (savedSnapRef.current===null) savedSnapRef.current = snapStr(); }, []); // baseline once
  const markClean = useCallback(()=>{ savedSnapRef.current = snapStr(); setDirty(false); try { localStorage.removeItem(AUTOSAVE_KEY); } catch {} }, [snapStr]);
  // detect edits and stash them locally (survives navigation / shutdown / crash)
  useEffect(()=>{ if (savedSnapRef.current===null) return; const cur = snapStr(); const d = cur!==savedSnapRef.current; setDirty(d);
    if (!d) return; const id = setTimeout(()=>{ try { localStorage.setItem(AUTOSAVE_KEY, JSON.stringify({ data:JSON.parse(cur), at:Date.now(), email:(session&&session.user.email)||"local" })); } catch {} }, 800);
    return ()=>clearTimeout(id); }, [snapStr, session]);
  // flush synchronously if the tab is closing
  useEffect(()=>{ const h=()=>{ if (savedSnapRef.current===null) return; const cur=snapStr(); if (cur!==savedSnapRef.current) { try { localStorage.setItem(AUTOSAVE_KEY, JSON.stringify({ data:JSON.parse(cur), at:Date.now(), email:(session&&session.user.email)||"local" })); } catch {} } };
    window.addEventListener("beforeunload", h); return ()=>window.removeEventListener("beforeunload", h); }, [snapStr, session]);
  // on first load, offer to restore any locally-stashed unsaved work
  useEffect(()=>{ try { const raw=localStorage.getItem(AUTOSAVE_KEY); if (raw) { const a=JSON.parse(raw); if (a && a.data && a.data.placed && a.data.placed.length) setRestore(a); } } catch {} }, []);
  // when signed in, also keep one cloud "Unsaved…" row up to date for cross-device recovery
  useEffect(()=>{ if (!supabase || !session) return;
    const iv = setInterval(async ()=>{ if (savedSnapRef.current===null) return; const cur=snapStr(); if (cur===savedSnapRef.current) return;
      const name = `Unsaved · ${session.user.email} · ${new Date().toLocaleString()}`;
      try { if (autoRowRef.current) { await supabase.from("schedules").update({ name, term:`${termStart}`, data:JSON.parse(cur) }).eq("id", autoRowRef.current); }
        else { const { data } = await supabase.from("schedules").insert({ user_id:session.user.id, name, term:`${termStart}`, status:"autosave", data:JSON.parse(cur) }).select("id").single(); if (data) autoRowRef.current=data.id; }
        refreshCloud(); } catch {} }, 60000);
    return ()=>clearInterval(iv); }, [supabase, session, snapStr, termStart, refreshCloud]);
  const applySnapshot = (s)=>{ if (!s || !s.placed) return; setPlaced(clone(s.placed)); setUnplaced(clone(s.unplaced||[]));
    if (s.courses) setCourses(clone(s.courses)); if (s.teachers) setTeachers(clone(s.teachers));
    if (s.termStart) setTermStart(s.termStart); if (s.roomPri) setRoomPri(clone(s.roomPri)); if (s.rules) setRules(clone(s.rules)); };
  const cloudSave = async (name, status)=>{ if (!supabase || !session) return;
    try {
      const { error } = await supabase.from("schedules").insert({ user_id:session.user.id, name:name||("Schedule "+new Date().toLocaleString()), term:`${termStart}`, status, data:cloudSnapshot() });
      if (error) showToast("conflict", `${t("cloudErr")}: ${error.message}`);
      else { showToast("valid", t("cloudSaved")); markClean();
        if (autoRowRef.current) { try { await supabase.from("schedules").delete().eq("id", autoRowRef.current); } catch {} autoRowRef.current=null; }
        refreshCloud(); }
    } catch (e) { showToast("conflict", `${t("cloudErr")}: ${e.message||e}`); }
  };
  const handleSave = ()=>{ // header Save: open the name + status dialog when signed in, else prompt sign-in
    if (session) { setSaveName(`${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString([], {hour:"2-digit", minute:"2-digit"})}`); setSaveStatus("checkpoint"); setShowSaveDialog(true); }
    else { setAuthMsg(""); setShowAuth(true); showToast("valid", t("signInToSave")); }
  };
  const cloudLoad = async (id)=>{ if (!supabase) return;
    const { data, error } = await supabase.from("schedules").select("data").eq("id",id).single();
    if (error || !data) { showToast("conflict", t("cloudErr")); return; }
    applySnapshot(data.data); setShowCloud(false); showToast("valid", t("cloudLoaded"));
    setTimeout(markClean, 0); // loaded state is the new clean baseline
  };
  const cloudDelete = async (id)=>{ if (!supabase) return; await supabase.from("schedules").delete().eq("id",id); refreshCloud(); };
  const cloudRename = async (id, cur)=>{ if (!supabase) return; const name = window.prompt(t("renameVersion"), cur); if (name==null || !name.trim()) return; await supabase.from("schedules").update({ name:name.trim() }).eq("id",id); refreshCloud(); };

  const visible = useMemo(()=>{
    let list = placed.filter((s)=>phaseVisible(s.phase, semView));
    if (view==="cohort") return list.filter((s)=>s.cohorts.includes(filterId));
    if (view==="instructor") return list.filter((s)=>s.ins===filterId);
    if (view==="room") return list.filter((s)=>s.room===filterId);
    return list;
  }, [placed, view, filterId, semView]);
  const cellSessions = (d,p,colId) => visible.filter((s)=> s.day===d && s.period===p && (colId ? s.cohorts.includes(colId) : true));

  const filterOptions =
    view==="cohort" ? COHORTS.map((c)=>({id:c.id,label:c.id})) :
    view==="instructor" ? teachers.map((i)=>({id:i.id,label:i.name})) :
    view==="room" ? ROOMS.map((r)=>({id:r.id,label:`${r.id} · ${r.type}`})) : [];
  useEffect(()=>{ if (view!=="master" && !filterOptions.find((o)=>o.id===filterId)) setFilterId(filterOptions[0]?.id); }, [view]); // eslint-disable-line

  const counts = useMemo(()=>({ cohorts:COHORTS.length, rooms:ROOMS.length, instructors:teachers.length, courses:courses.length, sessions:placed.length+unplaced.length }), [placed, unplaced, courses, teachers]);
  const dayOffCount = teachers.filter((tt)=>tt.off).length;
  const rating = useMemo(()=> (view==="instructor" && filterId) ? teacherRating(filterId, placed) : null, [view, filterId, placed]);
  const teacherRanks = useMemo(()=> teachers.map((tt)=>({ id:tt.id, name:tt.name, r:teacherRating(tt.id, placed) })).filter((x)=>x.r.total>0).sort((a,b)=>a.r.overall-b.r.overall), [teachers, placed]);

  const cellHandlers = { view, colorBy, lang, courses, teachers, onCard:setSelected, onLock:toggleLock, setDragId, half:semView };
  const dropProps = (d, p, cohort) => ({
    onDragOver:(e)=>{ if (dragId){ e.preventDefault(); setHoverCell(key(d,p)); } },
    onDrop:()=>{ if (dragId){ moveSession(dragId, d, p, cohort); setDragId(null); setHoverCell(null); } },
  });

  return (
    <div style={{ background:PAPER, color:INK, minHeight:"100vh", fontFamily:"'Segoe UI', system-ui, sans-serif" }}>
      <style>{`@media print {
          @page { size: A4 landscape; margin: 6mm; }
          html, body { background:#fff !important; }
          .no-print { display:none !important; }
          .print-full { width:100% !important; overflow:visible !important; }
          .print-fit { font-size:7.5px !important; }
          .print-fit table { width:100% !important; table-layout:fixed; }
          .print-fit td, .print-fit th { padding:2px 3px !important; min-width:0 !important; }
          .print-fit .rounded-md, .print-fit .rounded-lg { box-shadow:none !important; }
          .print-title { display:block !important; }
          tr, td, th, .print-fit > * { break-inside:avoid; }
          * { -webkit-print-color-adjust:exact !important; print-color-adjust:exact !important; }
        }
        .print-title { display:none; }
        *::-webkit-scrollbar { height:8px; width:8px; } *::-webkit-scrollbar-thumb { background:#cfcabb; border-radius:8px; }`}</style>
      <div className="print-title" style={{ padding:"0 0 6px", fontWeight:700, fontSize:14 }}>
        Мандах Их Сургууль — {t("title")} · {semView==="h1"?t("firstHalf"):t("secondHalf")}
      </div>

      {/* Header */}
      <header className="no-print sticky top-0 z-20" style={{ background:INK, color:PAPER }}>
        <div className="px-5 py-3 flex items-center gap-4 flex-wrap">
          <div className="flex items-center gap-2.5">
            <div className="grid place-items-center rounded-md" style={{ width:34, height:34, background:"#f7f6f2", color:INK }}><Calendar size={19} strokeWidth={2.2}/></div>
            <div><div className="font-bold text-[15px] tracking-tight">{t("title")}</div><div className="text-[11px] opacity-60">Мандах Их Сургууль · {t("subtitle")} <span className="opacity-70" title={t("buildNote")}>· build {BUILD_ID}</span></div></div>
          </div>
          <div className="flex-1"/>
          <div className="flex items-center rounded-md overflow-hidden text-[12px] font-semibold" style={{ border:"1px solid #ffffff33" }}>
            {["en","mn"].map((lc)=>(<button key={lc} onClick={()=>setLang(lc)} className="px-2.5 py-1 transition-colors" style={lang===lc ? { background:PAPER, color:INK } : { color:PAPER }}>{lc==="en"?"EN":"МН"}</button>))}
          </div>
          {score && screen==="timetable" && (
            <div className="flex items-center gap-2 rounded-lg px-3 py-1.5" style={{ background:"#ffffff14" }}>
              <Sparkles size={15} className="opacity-80"/><span className="text-[11px] opacity-70">{t("quality")}</span>
              <span className="text-lg font-bold leading-none">{score.overall}</span><span className="text-[11px] opacity-50 -ml-0.5">/100</span>
            </div>
          )}
          <div className="flex items-center gap-1.5">
            <button onClick={()=>runGenerate("all")} className="flex items-center gap-1.5 rounded-md px-3 py-1.5 text-[12px] font-semibold" style={{ background:"#f7f6f2", color:INK }}><Play size={13} strokeWidth={2.6}/> {t("generate")}</button>
            <button onClick={()=>runGenerate(hasHalves?"half":"all")} title={hasHalves?t("regenHalfHint"):undefined} className="flex items-center gap-1.5 rounded-md px-2.5 py-1.5 text-[12px]" style={{ background:"#ffffff1a" }}><RotateCw size={13}/> <span className="hidden lg:inline">{hasHalves ? `${t("regenerate")} · ${semView==="h1"?t("firstHalfShort"):t("secondHalfShort")}` : t("regenerate")}</span></button>
            <button onClick={handleSave} title={session?(dirty?t("unsavedChanges"):t("saveToCloud")):t("saveVersion")} className="relative flex items-center gap-1.5 rounded-md px-2.5 py-1.5 text-[12px]" style={{ background: session ? "#16a34a33" : "#ffffff1a" }}><Save size={13}/> <span className="hidden lg:inline">{t("save")}</span>{dirty && <span className="absolute -top-0.5 -right-0.5 rounded-full" style={{ width:8, height:8, background:"#f59e0b", border:"1.5px solid #26324a" }}/>}</button>
            <button onClick={()=>setShowRules(true)} className="flex items-center gap-1.5 rounded-md px-2.5 py-1.5 text-[12px]" style={{ background:"#ffffff1a" }}><Layers size={13}/> <span className="hidden lg:inline">{t("rulesBtn")}</span></button>
            {supabase && (session ? (
              <>
              <button onClick={()=>{ setShowCloud(true); refreshCloud(); }} className="flex items-center gap-1.5 rounded-md px-2.5 py-1.5 text-[12px]" style={{ background:"#16a34a33" }} title={session.user.email}><History size={13}/> <span className="hidden lg:inline">{t("mySchedules")}</span></button>
              {isAdmin && <button onClick={()=>{ setShowAdmin(true); refreshAdmin(); refreshUsers(); }} className="flex items-center gap-1.5 rounded-md px-2.5 py-1.5 text-[12px]" style={{ background:"#7c3aed44" }} title={t("adminHint")}><Gauge size={13}/> <span className="hidden lg:inline">{t("adminBtn")}</span></button>}
              </>
            ) : (
              <button onClick={()=>{ setAuthMsg(""); setShowAuth(true); }} className="flex items-center gap-1.5 rounded-md px-2.5 py-1.5 text-[12px]" style={{ background:"#ffffff1a" }}><User size={13}/> <span className="hidden lg:inline">{t("signIn")}</span></button>
            ))}
            <button onClick={()=>exportCSV(placed, lang, courses, teachers, semView)} className="rounded-md px-2 py-1.5" style={{ background:"#ffffff1a" }} title="Excel / CSV"><Download size={14}/></button>
            <button onClick={()=>window.print()} className="rounded-md px-2 py-1.5" style={{ background:"#ffffff1a" }} title="PDF"><Printer size={14}/></button>
          </div>
        </div>
        <div className="px-5 pb-2.5 flex items-center gap-2 flex-wrap">
          <div className="flex items-center rounded-full overflow-hidden text-[12px]" style={{ border:"1px solid #ffffff26" }}>
            {[{id:"timetable",label:t("timetable"),icon:LayoutGrid},{id:"data",label:t("courseData"),icon:Table2}].map((sc)=>(
              <button key={sc.id} onClick={()=>setScreen(sc.id)} className="flex items-center gap-1.5 px-3 py-1 transition-colors" style={screen===sc.id ? { background:PAPER, color:INK, fontWeight:600 } : { color:PAPER }}><sc.icon size={13}/> {sc.label}</button>
            ))}
          </div>
          {screen==="timetable" && <>
            <span className="opacity-20">|</span>
            {[{ id:"master", label:t("master"), icon:LayoutGrid }, { id:"cohort", label:t("byCohort"), icon:Users }, { id:"instructor", label:t("byInstructor"), icon:User }, { id:"room", label:t("byRoom"), icon:DoorOpen }].map((v)=>(
              <button key={v.id} onClick={()=>setView(v.id)} className="flex items-center gap-1.5 rounded-full px-3 py-1 text-[12px] transition-colors" style={view===v.id ? { background:PAPER, color:INK, fontWeight:600 } : { background:"#ffffff14", color:PAPER }}><v.icon size={13}/> {v.label}</button>
            ))}
            {view!=="master" && (<select value={filterId} onChange={(e)=>setFilterId(e.target.value)} className="ml-1 rounded-full px-3 py-1 text-[12px]" style={{ background:PAPER, color:INK }}>{filterOptions.map((o)=><option key={o.id} value={o.id}>{o.label}</option>)}</select>)}
            <div className="flex-1"/>
            {hasHalves && (
              <div className="flex items-center rounded-full overflow-hidden text-[12px]" style={{ border:"1px solid #ffffff26" }} title={`${t("weeks1")} / ${t("weeks2")}`}>
                {[{id:"h1",label:t("firstHalfShort")},{id:"h2",label:t("secondHalfShort")}].map((h)=>(<button key={h.id} onClick={()=>setSemView(h.id)} className="px-3 py-1 transition-colors" style={semView===h.id ? { background:PAPER, color:INK, fontWeight:600 } : { color:PAPER }}>{h.label}</button>))}
              </div>
            )}
            {hasHalves && (
              <button onClick={lockHalf} title={t("lockHalfHint")} className="flex items-center gap-1.5 rounded-md px-2.5 py-1 text-[12px]" style={halfAllLocked ? { background:"#b4530926", color:"#fed7aa" } : { background:"#ffffff1a" }}>
                {halfAllLocked ? <Unlock size={12}/> : <Lock size={12}/>} <span className="hidden lg:inline">{halfAllLocked ? t("unlockHalf") : t("lockHalf")} {semView==="h1"?t("firstHalfShort"):t("secondHalfShort")}</span>
              </button>
            )}
            <div className="flex items-center gap-1.5" title={t("colorByHint")}>
              <span className="text-[11px] opacity-60">{t("colorBy")}</span>
              <div className="flex items-center rounded-full overflow-hidden text-[12px]" style={{ border:"1px solid #ffffff26" }}>
                {[{id:"cohort",label:t("byCohortC")},{id:"type",label:t("byTypeC")}].map((m)=>(<button key={m.id} onClick={()=>setColorBy(m.id)} className="px-2.5 py-1 transition-colors" style={colorBy===m.id ? { background:PAPER, color:INK, fontWeight:600 } : { color:PAPER }}>{m.label}</button>))}
              </div>
            </div>
            {colorBy==="type" && (
              <div className="flex items-center gap-2 rounded-full px-2.5 py-1 text-[10px]" style={{ background:"#ffffff14" }}>
                {["L","S","B","Lab"].map((tp)=>(<span key={tp} className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm" style={{ background:COMP_COLORS[tp] }}/>{compTag(lang,tp)}</span>))}
              </div>
            )}
            <button onClick={()=>setShowData((s)=>!s)} className="flex items-center gap-1.5 text-[12px] opacity-80 hover:opacity-100"><Info size={13}/> {t("importSummary")}</button>
          </>}
        </div>
      </header>

      {/* Auth modal */}
      {showAuth && <AuthModal lang={lang} msg={authMsg} onClose={()=>setShowAuth(false)} onAuth={doAuth} />}
      {/* Cloud schedules modal */}
      {showCloud && session && (
        <CloudModal lang={lang} email={session.user.email} list={cloudList} termStart={termStart}
          onClose={()=>setShowCloud(false)} onSave={cloudSave} onLoad={cloudLoad} onDelete={cloudDelete} onSignOut={signOut} />
      )}
      {showAdmin && isAdmin && (
        <AdminModal lang={lang} list={adminList} users={adminUsers} selfId={session?.user?.id} onClose={()=>setShowAdmin(false)} onLoad={(id)=>{ cloudLoad(id); setShowAdmin(false); }} onRefresh={refreshAdmin} onRefreshUsers={refreshUsers} onSetAdmin={setAdminFor} />
      )}
      {restore && (
        <div className="no-print fixed bottom-5 left-1/2 -translate-x-1/2 z-50 flex items-center gap-3 rounded-lg px-4 py-3 shadow-lg" style={{ background:"#26324a", color:"#fff", maxWidth:"92vw" }}>
          <RotateCw size={16} className="shrink-0"/>
          <div className="text-[12px]"><div className="font-semibold">{t("restoreTitle")}</div><div className="opacity-70">{t("restoreFrom")} {new Date(restore.at).toLocaleString()}{restore.email && restore.email!=="local" ? " · "+restore.email : ""}</div></div>
          <button onClick={()=>{ applySnapshot(restore.data); setRestore(null); showToast("valid", t("restored")); }} className="rounded-md px-3 py-1.5 text-[12px] font-semibold shrink-0" style={{ background:"#fff", color:"#26324a" }}>{t("restoreBtn")}</button>
          <button onClick={()=>{ try{ localStorage.removeItem(AUTOSAVE_KEY); }catch{} setRestore(null); }} className="rounded-md px-2.5 py-1.5 text-[12px] shrink-0" style={{ background:"#ffffff22" }}>{t("discardBtn")}</button>
        </div>
      )}
      {showSaveDialog && session && (
        <div className="no-print fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background:"#1a2438aa" }} onClick={()=>setShowSaveDialog(false)}>
          <div className="rounded-xl max-w-sm w-full" style={{ background:PAPER }} onClick={(e)=>e.stopPropagation()}>
            <div className="flex items-center justify-between px-5 py-3" style={{ background:"#16a34a", color:PAPER }}>
              <div className="flex items-center gap-2 font-semibold text-[14px]"><Save size={15}/> {t("saveVersionTitle")}</div>
              <button onClick={()=>setShowSaveDialog(false)} className="rounded-md p-1" style={{ background:"#ffffff1a" }}><X size={16}/></button>
            </div>
            <div className="p-5 space-y-4">
              <div>
                <label className="text-[11px] font-semibold opacity-70">{t("versionName")}</label>
                <input value={saveName} onChange={(e)=>setSaveName(e.target.value)} autoFocus className="mt-1 w-full rounded border px-2.5 py-2 text-[13px]" style={{ borderColor:LINE }} placeholder={t("versionNamePh")} />
              </div>
              <div>
                <label className="text-[11px] font-semibold opacity-70">{t("versionStatus")}</label>
                <div className="mt-1.5 grid grid-cols-2 gap-2">
                  {[["checkpoint", t("stInProgress"), "#16a34a"],["final", t("stFinished"), "#7c3aed"]].map(([id,label,col])=>(
                    <button key={id} onClick={()=>setSaveStatus(id)} className="rounded-lg px-3 py-2 text-[12px] font-medium border-2" style={ saveStatus===id ? { borderColor:col, background:`${col}12`, color:col } : { borderColor:LINE, background:"#fff", color:INK }}>{label}</button>
                  ))}
                </div>
                <div className="text-[10px] opacity-55 mt-1.5">{t("saveOwnerNote").replace("{email}", session.user.email)}</div>
              </div>
              <button onClick={()=>{ cloudSave(saveName.trim()||new Date().toLocaleString(), saveStatus); setShowSaveDialog(false); }} className="w-full rounded-lg py-2.5 text-[13px] font-semibold flex items-center justify-center gap-2" style={{ background:INK, color:PAPER }}><Save size={14}/> {t("saveVersionBtn")}</button>
            </div>
          </div>
        </div>
      )}

      {/* Rules panel */}
      {showRules && (
        <div className="no-print fixed inset-0 z-40 flex items-center justify-center p-4" style={{ background:"#1a2438aa" }} onClick={()=>setShowRules(false)}>
          <div className="rounded-xl max-w-2xl w-full max-h-[85vh] overflow-auto" style={{ background:PAPER }} onClick={(e)=>e.stopPropagation()}>
            <div className="flex items-center justify-between px-5 py-3 sticky top-0" style={{ background:INK, color:PAPER }}>
              <div className="flex items-center gap-2 font-semibold text-[14px]"><Layers size={15}/> {t("rulesTitle")}</div>
              <button onClick={()=>setShowRules(false)} className="rounded-md p-1" style={{ background:"#ffffff1a" }}><X size={16}/></button>
            </div>
            <div className="p-5 space-y-5">
              <div>
                <div className="flex items-center gap-1.5 text-[12px] font-semibold mb-2" style={{ color:INK }}><Lock size={13}/> {t("hardRules")}</div>
                <ul className="space-y-1.5">
                  {RULES_HARD.map((r,i)=>(<li key={i} className="flex items-start gap-2 text-[13px]"><Check size={15} className="mt-0.5 shrink-0" style={{ color:"#15803d" }}/><span>{r[lang]}</span></li>))}
                </ul>
              </div>
              <div style={{ borderTop:`1px solid ${LINE}` }} className="pt-4">
                <div className="flex items-center gap-1.5 text-[12px] font-semibold mb-1" style={{ color:INK }}><Sparkles size={13}/> {t("softRules")}</div>
                <div className="text-[11px] opacity-60 mb-2.5">{t("rulesHint")}</div>
                <div className="space-y-1.5">
                  {RULES_SOFT.map((r)=>(
                    <label key={r.id} className="flex items-center justify-between gap-3 rounded-lg px-3 py-2 cursor-pointer" style={{ background:"#fff", border:`1px solid ${LINE}` }}>
                      <span className="text-[13px]">{r[lang]}</span>
                      <button type="button" onClick={()=>setRules((prev)=>({ ...prev, [r.id]:!prev[r.id] }))} className="relative rounded-full transition-colors shrink-0" style={{ width:38, height:22, background: rules[r.id] ? INK : "#cfcabb" }}>
                        <span className="absolute top-0.5 rounded-full bg-white transition-all" style={{ width:18, height:18, left: rules[r.id] ? 18 : 2 }}/>
                      </button>
                    </label>
                  ))}
                </div>
              </div>
              <button onClick={()=>{ setShowRules(false); runGenerate("all"); }} className="w-full rounded-lg py-2.5 text-[13px] font-semibold flex items-center justify-center gap-2" style={{ background:INK, color:PAPER }}><RotateCw size={14}/> {t("applyRegen")}</button>
            </div>
          </div>
        </div>
      )}

      {/* Import summary */}
      {showData && screen==="timetable" && (
        <div className="no-print px-5 py-3 border-b" style={{ background:"#fff", borderColor:LINE }}>
          <div className="text-[11px] opacity-60 mb-1.5">{t("importWhat")}</div>
          <div className="text-[12px] flex flex-wrap gap-x-6 gap-y-2 items-center">
            <span className="font-semibold">{counts.cohorts} {t("cohortsN")}</span><span>{counts.courses} {t("coursesN")}</span>
            <span>{counts.instructors} {t("instructorsN")}</span><span>{counts.rooms} {t("roomsN")}</span><span>{counts.sessions} {t("sessionsN")}</span>
            {dayOffCount>0 && <span className="flex items-center gap-1" style={{ color:"#b45309" }}><AlertTriangle size={13}/> {dayOffCount} {t("dayOffNote")}</span>}
            {unplaced.length>0 && <span className="flex items-center gap-1" style={{ color:"#be123c" }}><AlertTriangle size={13}/> {unplaced.length} {t("attentionN")}</span>}
          </div>
        </div>
      )}

      {/* Body */}
      {screen==="data" ? (
        <div className="p-4"><CourseEditor lang={lang} courses={courses} setCourses={setCourses} teachers={teachers} setTeachers={setTeachers} termStart={termStart} setTermStart={setTermStart} roomPri={roomPri} setRoomPri={setRoomPri} onApply={()=>{ runGenerate("fresh"); setScreen("timetable"); }} /></div>
      ) : (
        <div className="flex gap-4 p-4 items-start">
          <main className="flex-1 min-w-0 print-full print-fit">
            <div className="rounded-xl overflow-hidden border" style={{ borderColor:LINE, background:"#fff" }}>
              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-left" style={{ minWidth: view==="master" ? 980 : 680 }}>
                  <thead>
                    <tr style={{ background:"#fbfaf6" }}>
                      {view==="master" ? (<>
                        <th className="sticky left-0 z-10 px-2 py-2 text-[11px] font-semibold" style={{ background:"#fbfaf6", borderBottom:`1px solid ${LINE}`, borderRight:`1px solid ${LINE}`, width:60 }}>{t("day")}</th>
                        <th className="px-2 py-2 text-[11px] font-semibold" style={{ background:"#fbfaf6", borderBottom:`1px solid ${LINE}`, borderRight:`1px solid ${LINE}`, width:96 }}>{t("period")}</th>
                        {COHORTS.map((c, idx)=>(<th key={c.id} className="px-2 py-2 text-[12px]" style={{ borderBottom:`2px solid ${THICK}`, borderLeft: (idx>0 && COHORTS[idx].year!==COHORTS[idx-1].year) ? `3px solid ${THICK}` : `1px solid ${LINE}` }}><div className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full" style={{ background:COHORT_COLORS[c.id] }}/><span className="font-bold">{c.id}</span><span className="opacity-50 text-[10px] whitespace-nowrap">{t("year")} {c.year} · {c.students}</span></div></th>))}
                      </>) : (<>
                        <th className="sticky left-0 z-10 px-2 py-2 text-[11px] font-semibold" style={{ background:"#fbfaf6", borderBottom:`1px solid ${LINE}`, borderRight:`1px solid ${LINE}`, width:74 }}>{t("period")}</th>
                        {DAYS.map((d)=>(<th key={d.id} className="px-2 py-2 text-[12px] font-bold" style={{ borderBottom:`1px solid ${LINE}` }}>{dayLabel(lang,d)}</th>))}
                      </>)}
                    </tr>
                  </thead>
                  <tbody>
                    {view==="master"
                      ? DAYS.map((d)=>PERIODS.map((p, pi)=>(
                          <tr key={d.id+p.id}>
                            {pi===0 && (<td rowSpan={PERIODS.length} className="sticky left-0 z-10 align-top px-2 py-2" style={{ background:"#fff", borderRight:`2px solid ${THICK}`, borderBottom:`3px solid ${THICK}`, width:60 }}><div className="text-[12px] font-bold leading-tight">{dayLabel(lang,d)}</div></td>)}
                            <td className="px-2 py-1.5 align-top" style={{ background:"#fff", borderRight:`1px solid ${LINE}`, borderBottom: pi===PERIODS.length-1 ? `3px solid ${THICK}` : `1px solid ${LINE}`, width:96 }}><div className="text-[11px] font-semibold">P{p.id}</div><div className="text-[10px] opacity-50 whitespace-nowrap">{p.label}</div></td>
                            {(()=>{ const out=[]; let i=0;
                              const rowEmpty = COHORTS.every((c)=>cellSessions(d.id, p.id, c.id).length===0);
                              const lastRow = pi===PERIODS.length-1;
                              while (i < COHORTS.length) {
                                const ci = COHORTS[i].id; const ss = cellSessions(d.id, p.id, ci);
                                const yearStart = i>0 && COHORTS[i].year!==COHORTS[i-1].year;
                                let span = 1;
                                if (ss.length) { const idset = ss.map((x)=>x.id).sort().join(",");
                                  while (i+span < COHORTS.length) { const nid = COHORTS[i+span].id; const ns = cellSessions(d.id, p.id, nid);
                                    if (ns.map((x)=>x.id).sort().join(",")===idset && ss.every((x)=>x.cohorts.includes(nid))) span++; else break; } }
                                const blocked = d.id==="tue" && p.id>=TUE_CUTOFF_PERIOD;
                                const hatched = blocked || rowEmpty; // gray out blocked slots AND fully-empty period rows
                                out.push(<td key={ci} colSpan={span} {...dropProps(d.id,p.id,ci)} className="align-top px-1.5 py-1.5" style={{ height:"46px", borderBottom: lastRow ? `3px solid ${THICK}` : `1px solid ${LINE}`, borderLeft: yearStart ? `3px solid ${THICK}` : `1px solid ${LINE}`, background: hatched ? "repeating-linear-gradient(45deg,#faf8f2,#faf8f2 6px,#f2efe6 6px,#f2efe6 12px)" : hoverCell===key(d.id,p.id) ? "#eef6f4" : "#fff", minWidth: span>1?undefined:140 }}><CellContent list={ss} {...cellHandlers}/></td>);
                                i += span;
                              }
                              return out;
                            })()}
                          </tr>
                        )))
                      : PERIODS.map((p)=>(
                          <tr key={p.id}>
                            <td className="sticky left-0 z-10 px-2 py-2 align-top" style={{ background:"#fff", borderRight:`1px solid ${LINE}`, borderBottom:`1px solid ${LINE}`, width:74 }}><div className="text-[11px] font-semibold">P{p.id}</div><div className="text-[10px] opacity-50 whitespace-nowrap">{p.label}</div></td>
                            {DAYS.map((d)=>{ const ss = cellSessions(d.id, p.id, null); const blocked = d.id==="tue" && p.id>=TUE_CUTOFF_PERIOD;
                              return (<td key={d.id} {...dropProps(d.id,p.id)} className="align-top px-1.5 py-1.5" style={{ height:"46px", borderBottom:`1px solid ${LINE}`, borderLeft:`1px solid ${LINE}`, background: blocked ? "repeating-linear-gradient(45deg,#faf8f2,#faf8f2 6px,#f2efe6 6px,#f2efe6 12px)" : hoverCell===key(d.id,p.id) ? "#eef6f4" : "#fff", minWidth:158 }}><CellContent list={ss} {...cellHandlers}/></td>);
                            })}
                          </tr>
                        ))}
                  </tbody>
                </table>
              </div>
            </div>
            <p className="no-print mt-2 text-[11px] opacity-50 flex items-center gap-1.5"><ChevronRight size={12}/> {t("dragHint")}</p>
          </main>

          {/* Right rail */}
          <aside className="no-print w-72 shrink-0 flex flex-col gap-4" style={{ maxWidth:288 }}>
            {/* teacher rating (instructor view) */}
            {view==="instructor" && rating && (
              <Panel title={t("teacherRating")} icon={Gauge}>
                <div className="text-[12px] font-semibold mb-1">{teacherName(teachers, filterId)}</div>
                <div className="flex items-baseline gap-1 mb-3"><span className="text-3xl font-bold">{rating.overall}</span><span className="opacity-40 text-sm">/ 100</span></div>
                <div className="space-y-2">
                  <Bar label={t("workload")} value={rating.workload}/>
                  <Bar label={t("breaksLbl")} value={rating.breaks}/>
                  <Bar label={t("consecutive")} value={rating.consecutive}/>
                </div>
                <div className="mt-3 pt-3 text-[11px] opacity-60 flex flex-wrap gap-x-3 gap-y-0.5" style={{ borderTop:`1px solid ${LINE}` }}>
                  <span>{rating.courseCount} {t("coursesN")}</span><span className="font-semibold opacity-90">{t("firstHalfShort")}: {rating.perHalf.h1} · {t("secondHalfShort")}: {rating.perHalf.h2} {t("periodsWk")}</span><span>{t("gaps")}: {rating.gaps}</span>
                </div>
                <div className="mt-3 pt-3" style={{ borderTop:`1px solid ${LINE}` }}>
                  <div className="text-[10px] font-semibold opacity-60 mb-1.5 uppercase tracking-wide">{t("allTeachers")}</div>
                  <div className="space-y-1">
                    {teacherRanks.map((x)=>{ const col = x.r.overall>=90?"#0d9488":x.r.overall>=75?"#b45309":"#be123c";
                      return (<button key={x.id} onClick={()=>setFilterId(x.id)} className="w-full flex items-center gap-2 text-[11px] rounded px-1.5 py-1" style={x.id===filterId?{background:"#fbfaf6"}:{}}>
                        <span className="w-1.5 h-1.5 rounded-full" style={{ background:col }}/>
                        <span className="flex-1 text-left truncate">{x.name}</span>
                        <span className="opacity-50 tabular-nums" title={t("periodsWk")}>{t("firstHalfShort").slice(0,1)}{x.r.perHalf.h1}·{t("secondHalfShort").slice(0,1)}{x.r.perHalf.h2}</span>
                        <span className="font-semibold" style={{ color:col }}>{x.r.overall}</span>
                      </button>);
                    })}
                  </div>
                </div>
              </Panel>
            )}

            {score && (
              <Panel title={`${t("scheduleQuality")} · ${semView==="h1"?t("firstHalfShort"):t("secondHalfShort")}`} icon={Sparkles}>
                <div className="flex items-baseline gap-1 mb-3"><span className="text-3xl font-bold">{score.overall}</span><span className="opacity-40 text-sm">/ 100</span></div>
                <div className="space-y-2">
                  <Bar label={t("hardConstraints")} value={score.parts.hard}/><Bar label={t("studentExp")} value={score.parts.student}/>
                  <Bar label={t("instructorPref")} value={score.parts.instructor}/><Bar label={t("dailyBalance")} value={score.parts.balance}/><Bar label={t("roomEff")} value={score.parts.room}/>
                </div>
                <div className="mt-3 pt-3 text-[11px] opacity-60 space-y-0.5" style={{ borderTop:`1px solid ${LINE}` }}>
                  <div>{t("gaps")}: {score.metrics.gaps}</div><div>{t("p4used")}: {score.metrics.p4}</div><div>{t("tueLate")}: {score.metrics.tueLate}</div>
                </div>
              </Panel>
            )}

            {/* candidate options */}
            {candidates.length>0 && (
              <Panel title={t("candidates")} icon={Layers}>
                <p className="text-[11px] opacity-55 mb-2">{t("pickOption")}</p>
                <div className="space-y-1.5">
                  {candidates.slice(0,6).map((c,i)=>{ const hs = scoreSchedule(c.placed, c.unplaced, teachers, semView); const col = hs.overall>=90?"#0d9488":hs.overall>=75?"#b45309":"#be123c"; const act = i===activeCand;
                    const hUn = c.unplaced.filter((s)=>phaseVisible(s.phase, semView)).length;
                    return (<button key={i} onClick={()=>selectCandidate(i)} className="w-full flex items-center gap-2 rounded-md px-2 py-1.5 text-[12px]" style={act?{ background:INK, color:PAPER }:{ background:"#fbfaf6" }}>
                      {act ? <Check size={13}/> : <span className="w-3.5"/>}
                      <span className="flex-1 text-left font-medium">{t("option")} {i+1}</span>
                      {hUn>0 && <span className="text-[10px]" style={{ color:act?"#ffb4a4":"#be123c" }}>{hUn}⚠</span>}
                      <span className="font-bold" style={act?{}:{ color:col }}>{hs.overall}</span>
                    </button>);
                  })}
                </div>
              </Panel>
            )}

            {overloaded.length>0 && (
              <Panel title={t("capacityTitle")} icon={AlertTriangle} accent="#b45309">
                <div className="space-y-1.5 text-[11px]">
                  <p className="opacity-70">{t("capacityDesc").replace("{n}", SLOTS_PER_HALF)}</p>
                  {overloaded.map((c)=>(<div key={c.id} className="rounded-md px-2 py-1.5" style={{ background:"#b4530911" }}>
                    <span className="font-semibold">{c.id}</span> — {c.h1>=SLOTS_PER_HALF-1 && <span>{t("firstHalf")}: {c.h1}/{SLOTS_PER_HALF}</span>}{c.h1>=SLOTS_PER_HALF-1 && c.h2>=SLOTS_PER_HALF-1 && " · "}{c.h2>=SLOTS_PER_HALF-1 && <span>{t("secondHalf")}: {c.h2}/{SLOTS_PER_HALF}</span>}
                  </div>))}
                </div>
              </Panel>
            )}
            {conflicts.length>0 && (
              <Panel title={`${t("conflictsTitle")} (${conflicts.length})`} icon={AlertTriangle} accent="#be123c">
                <div className="space-y-1.5 text-[11px]">
                  <p className="opacity-70">{t("conflictsDesc")}</p>
                  {conflicts.map((c,i)=>{ const ent = c.kind==="instructor" ? teacherName(teachers,c.ent) : c.ent; const names=[...new Set(c.sessions.map((s)=>courseName(courses,s.courseIdx)))];
                    return (<div key={i} className="rounded-md px-2 py-1.5" style={{ background:"#be123c12" }}>
                      <div className="font-semibold">{ent} · {dayLabel(lang,DAYS.find((d)=>d.id===c.day))} · P{c.period} <span className="opacity-60 font-normal">({t("cf_"+c.kind)})</span></div>
                      <div className="opacity-75">{names.join("  +  ")}</div>
                    </div>); })}
                </div>
              </Panel>
            )}
            {repeats.length>0 && (
              <Panel title={`${t("repeatsTitle")} (${repeats.length})`} icon={AlertTriangle} accent="#be123c">
                <div className="space-y-1.5 text-[11px]">
                  <p className="opacity-70">{t("repeatsDesc")}</p>
                  {repeats.map((r,i)=>(<div key={i} className="rounded-md px-2 py-1.5" style={{ background:"#be123c10" }}>
                    <div className="font-semibold">{courseName(courses,r.courseIdx)}</div>
                    <div className="opacity-70">{r.cohort} · {dayLabel(lang,DAYS.find((d)=>d.id===r.day))} · P{r.periods.join(", P")} · {teacherName(teachers,r.ins)}</div>
                  </div>))}
                </div>
              </Panel>
            )}
            {unplaced.length>0 && (
              <Panel title={`${t("needsAttention")} (${unplaced.length})`} icon={AlertTriangle} accent="#be123c">
                <div className="space-y-1.5 text-[11px]">
                  {unplaced.map((s)=>(<div key={s.id} className="rounded-md px-2 py-1.5" style={{ background:"#be123c10" }}><div className="font-semibold">{courseName(courses,s.courseIdx)} /{compTag(lang,s.type)}/</div><div className="opacity-70">{s.cohorts.join("+")} · {teacherName(teachers,s.ins)}</div></div>))}
                  <p className="opacity-60 pt-1">{t("unplacedHint")}</p>
                </div>
              </Panel>
            )}
            {selected && <Panel title={t("classDetail")} icon={Info} onClose={()=>setSelected(null)}><SelectedDetail s={selected} lang={lang} courses={courses} teachers={teachers} placed={placed} half={semView} onLock={()=>toggleLock(selected.id)} onSetRoom={setRoomManual}/></Panel>}
            <Panel title={t("versions")} icon={History}>
              {session ? (
                cloudList.length===0 ? <p className="text-[11px] opacity-55">{t("noVersions")}</p> : (
                  <div className="space-y-1.5">{[...cloudList].sort((a,b)=>new Date(a.updated_at)-new Date(b.updated_at)).map((v,i)=>(
                    <div key={v.id} className="flex items-center gap-1.5 rounded-md px-2 py-1.5 text-[11px]" style={{ background:"#fbfaf6" }}>
                      <span className="font-semibold shrink-0">V{i+1}</span>
                      <div className="flex-1 min-w-0"><div className="truncate">{v.name}</div><div className="opacity-45 text-[10px]">{new Date(v.updated_at).toLocaleString()}</div></div>
                      <button onClick={()=>cloudLoad(v.id)} title={t("loadBtn")} className="rounded px-2 py-0.5 font-medium shrink-0" style={{ background:INK, color:PAPER }}>{t("loadBtn")}</button>
                      <button onClick={()=>cloudRename(v.id, v.name)} title={t("renameVersion")} className="opacity-40 hover:opacity-100 shrink-0"><Info size={13}/></button>
                      <button onClick={()=>cloudDelete(v.id)} title={t("delete")} className="opacity-40 hover:opacity-100 shrink-0" style={{ color:"#be123c" }}><Trash2 size={13}/></button>
                    </div>))}</div>
                )
              ) : (
                <p className="text-[11px] opacity-55">{t("signInSaveVersions")}</p>
              )}
            </Panel>
          </aside>
        </div>
      )}

      {generating && (<div className="no-print fixed inset-0 z-40 grid place-items-center" style={{ background:"#1a2438cc", color:PAPER }}><div className="text-center"><Layers size={30} className="mx-auto mb-3 animate-pulse"/><div className="text-sm font-semibold">{t(genStep)}</div><div className="text-[11px] opacity-60 mt-1">{t("solverNote")}</div></div></div>)}
      {toast && (<div className="no-print fixed bottom-5 left-1/2 -translate-x-1/2 z-50 flex items-center gap-2 rounded-lg px-4 py-2.5 text-[12px] font-medium shadow-lg" style={{ background: toast.status==="conflict" ? "#be123c" : toast.status==="warning" ? "#b45309" : "#0d9488", color:"#fff" }}>{toast.status==="conflict" ? <X size={15}/> : toast.status==="warning" ? <AlertTriangle size={15}/> : <CheckCircle2 size={15}/>}{toast.text}</div>)}
    </div>
  );
}

const key = (d,p) => `${d}|${p}`;

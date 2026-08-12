import type { AppData, Placement } from "./types";
import {
  buildMonthCalendar,
  getCycleLength,
  getEntriesForDate,
  isSunday,
  parseISODate,
  weekdayLabel,
} from "./schedule";
import { triggerDownload } from "./csv";

const COLORS = [
  "FF6366F1",
  "FF10B981",
  "FFF59E0B",
  "FFEF4444",
  "FF8B5CF6",
  "FF06B6D4",
  "FFEC4899",
  "FF84CC16",
  "FFF97316",
  "FF14B8A6",
];

function xmlEscape(value: string | number | null | undefined): string {
  if (value === null || value === undefined) return "";
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function cell(
  value: string | number | null | undefined,
  options: { type?: string; styleId?: string } = {},
): string {
  const type = options.type ?? (typeof value === "number" ? "Number" : "String");
  const styleAttr = options.styleId ? ` ss:StyleID="${options.styleId}"` : "";
  const safeValue = value ?? "";
  return `<Cell${styleAttr}><Data ss:Type="${type}">${xmlEscape(safeValue)}</Data></Cell>`;
}

function row(cells: string[], options: { height?: number } = {}): string {
  const heightAttr = options.height ? ` ss:Height="${options.height}"` : "";
  return `<Row${heightAttr}>${cells.join("")}</Row>`;
}

function buildStyles(): string {
  const styles: string[] = [];

  styles.push(`<Style ss:ID="Default" ss:Name="Normal">
    <Font ss:FontName="Calibri" ss:Size="11"/>
    <Alignment ss:Vertical="Center" ss:WrapText="1"/>
  </Style>`);

  styles.push(`<Style ss:ID="Title">
    <Font ss:FontName="Calibri" ss:Size="18" ss:Bold="1" ss:Color="FFFFFFFF"/>
    <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
    <Interior ss:Color="FF1F2937" ss:Pattern="Solid"/>
  </Style>`);

  styles.push(`<Style ss:ID="Subtitle">
    <Font ss:FontName="Calibri" ss:Size="12" ss:Bold="1" ss:Color="FFFFFFFF"/>
    <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
    <Interior ss:Color="FF374151" ss:Pattern="Solid"/>
  </Style>`);

  styles.push(`<Style ss:ID="Header">
    <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="FFFFFFFF"/>
    <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
    <Interior ss:Color="FF4B5563" ss:Pattern="Solid"/>
    <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FF1F2937"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FF1F2937"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FF1F2937"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FF1F2937"/>
    </Borders>
  </Style>`);

  styles.push(`<Style ss:ID="Cell">
    <Alignment ss:Vertical="Center" ss:WrapText="1"/>
    <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FFD1D5DB"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FFD1D5DB"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FFD1D5DB"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FFD1D5DB"/>
    </Borders>
  </Style>`);

  styles.push(`<Style ss:ID="CellCenter">
    <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
    <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FFD1D5DB"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FFD1D5DB"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FFD1D5DB"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FFD1D5DB"/>
    </Borders>
  </Style>`);

  styles.push(`<Style ss:ID="Sunday">
    <Font ss:FontName="Calibri" ss:Size="10" ss:Italic="1" ss:Color="FF9CA3AF"/>
    <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
    <Interior ss:Color="FFF3F4F6" ss:Pattern="Solid"/>
  </Style>`);

  COLORS.forEach((color, idx) => {
    styles.push(`<Style ss:ID="GroupHeader${idx}">
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="FFFFFFFF"/>
      <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
      <Interior ss:Color="${color}" ss:Pattern="Solid"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FF1F2937"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FF1F2937"/>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FF1F2937"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="FF1F2937"/>
      </Borders>
    </Style>`);
  });

  return `<Styles>${styles.join("")}</Styles>`;
}

interface SheetInput {
  name: string;
  rows: string[];
  columnWidths?: number[];
}

function buildSheet(sheet: SheetInput): string {
  const columns = sheet.columnWidths
    ? `<Columns>${sheet.columnWidths
        .map((w) => `<Column ss:Width="${w}"/>`)
        .join("")}</Columns>`
    : "";

  return `<Worksheet ss:Name="${xmlEscape(sheet.name)}">
    ${columns}
    <Table>${sheet.rows.join("")}</Table>
  </Worksheet>`;
}

function buildSettingsSheet(data: AppData): string {
  const rows: string[] = [];
  rows.push(
    row(
      [
        cell("Настройки и данные", { styleId: "Title" }),
        cell("", { styleId: "Title" }),
      ],
      { height: 30 },
    ),
  );
  rows.push(row([cell("", { styleId: "Cell" }), cell("", { styleId: "Cell" })]));

  const writeSection = (title: string, items: [string, string][]) => {
    rows.push(
      row(
        [
          cell(title, { styleId: "Subtitle" }),
          cell("", { styleId: "Subtitle" }),
        ],
        { height: 22 },
      ),
    );
    items.forEach(([k, v]) => {
      rows.push(row([cell(k, { styleId: "Cell" }), cell(v, { styleId: "Cell" })]));
    });
    rows.push(row([cell("", { styleId: "Cell" }), cell("", { styleId: "Cell" })]));
  };

  const cycleDays = getCycleLength(data);
  const defaultSlot = data.settings.timeSlots[0];
  const defaultRange = defaultSlot
    ? `${defaultSlot.start} – ${defaultSlot.end}`
    : "—";
  writeSection("Общие настройки", [
    ["Дата начала обучения", data.settings.startDate],
    ["Длительность цикла (учебных дней)", String(cycleDays)],
    ["Время по умолчанию", defaultRange],
    ["Тем в курсе", String(data.topics.length)],
    [
      "Все временные интервалы",
      data.settings.timeSlots.length
        ? data.settings.timeSlots
            .map(
              (s, i) =>
                `${i + 1}. ${s.start} – ${s.end}${i === 0 ? " (по умолчанию)" : ""}`,
            )
            .join("\n")
        : "—",
    ],
  ]);

  writeSection(
    "Преподаватели",
    data.teachers.length
      ? data.teachers.map((t, i) => [`${i + 1}. ${t.name}`, t.id])
      : [["(пусто)", ""]],
  );

  writeSection(
    "Места обучения",
    data.locations.length
      ? data.locations.map((l, i) => [`${i + 1}. ${l.name}`, l.id])
      : [["(пусто)", ""]],
  );

  writeSection(
    "Темы занятий",
    data.topics.length
      ? data.topics.map((t, i) => [`${i + 1}. ${t.name}`, t.id])
      : [["(пусто)", ""]],
  );

  writeSection(
    "Группы",
    data.groups.length
      ? data.groups.map((g, i) => [`${i + 1}. ${g.name}`, g.id])
      : [["(пусто)", ""]],
  );

  writeSection(
    "Размещения (placements)",
    data.placements.length
      ? data.placements.map((p, i) => {
          const group = data.groups.find((g) => g.id === p.groupId)?.name ?? "?";
          const teachers = p.teacherIds
            .map((id) => data.teachers.find((t) => t.id === id)?.name ?? "")
            .filter(Boolean)
            .join(", ");
          const location =
            data.locations.find((l) => l.id === p.locationId)?.name ?? "—";
          const timeLabel =
            p.customStartTime && p.customEndTime
              ? `${p.customStartTime} – ${p.customEndTime} (своё)`
              : p.timeSlotId
                ? `${(() => {
                    const slot = data.settings.timeSlots.find(
                      (s) => s.id === p.timeSlotId,
                    );
                    return slot ? `${slot.start} – ${slot.end}` : defaultRange;
                  })()}`
                : defaultRange;
          return [
            `${i + 1}. ${group}`,
            `Старт: ${p.startDate} | Время: ${timeLabel} | Преподаватели: ${teachers || "—"} | Место: ${location}`,
          ];
        })
      : [["(пусто)", ""]],
  );

  return buildSheet({
    name: "Настройки",
    rows,
    columnWidths: [240, 480],
  });
}

function buildOverviewSheet(data: AppData): string {
  const rows: string[] = [];

  rows.push(
    row(
      [
        cell(`Расписание: с ${data.settings.startDate}`, { styleId: "Title" }),
        cell("", { styleId: "Title" }),
        cell("", { styleId: "Title" }),
      ],
      { height: 30 },
    ),
  );

  rows.push(
    row(
      [
        cell(`Длительность цикла: ${cycleDays} уч. дн. (по числу тем)`, {
          styleId: "Subtitle",
        }),
        cell("", { styleId: "Subtitle" }),
        cell("", { styleId: "Subtitle" }),
      ],
      { height: 22 },
    ),
  );

  rows.push(row([cell("", { styleId: "Cell" })]));

  rows.push(
    row(
      [
        cell("Раздел", { styleId: "Header" }),
        cell("Название", { styleId: "Header" }),
        cell("Доп. информация", { styleId: "Header" }),
      ],
      { height: 22 },
    ),
  );

  data.teachers.forEach((t) => {
    rows.push(
      row([
        cell("Преподаватель", { styleId: "Cell" }),
        cell(t.name, { styleId: "Cell" }),
        cell("", { styleId: "Cell" }),
      ]),
    );
  });

  data.locations.forEach((l) => {
    rows.push(
      row([
        cell("Место обучения", { styleId: "Cell" }),
        cell(l.name, { styleId: "Cell" }),
        cell("", { styleId: "Cell" }),
      ]),
    );
  });

  data.topics.forEach((t) => {
    rows.push(
      row([
        cell("Тема занятия", { styleId: "Cell" }),
        cell(t.name, { styleId: "Cell" }),
        cell("", { styleId: "Cell" }),
      ]),
    );
  });

  data.groups.forEach((g) => {
    rows.push(
      row([
        cell("Группа", { styleId: "Cell" }),
        cell(g.name, { styleId: "Cell" }),
        cell("", { styleId: "Cell" }),
      ]),
    );
  });

  return buildSheet({ name: "Обзор", rows, columnWidths: [140, 220, 380] });
}

function buildGroupScheduleSheet(
  group: AppData["groups"][number],
  data: AppData,
  groupIndex: number,
): string {
  const rows: string[] = [];
  const colorStyleId = `GroupHeader${groupIndex % COLORS.length}`;
  const placements = data.placements.filter((p) => p.groupId === group.id);

  rows.push(
    row(
      [
        cell(`Группа: ${group.name}`, { styleId: colorStyleId }),
        cell("", { styleId: colorStyleId }),
        cell("", { styleId: colorStyleId }),
        cell("", { styleId: colorStyleId }),
        cell("", { styleId: colorStyleId }),
      ],
      { height: 28 },
    ),
  );

  rows.push(
    row(
      [
        cell(
          `Размещений: ${placements.length} | Тем в цикле: ${data.topics.length} | Длительность цикла: ${cycleDays} уч. дн.`,
          { styleId: "Cell" },
        ),
        cell("", { styleId: "Cell" }),
        cell("", { styleId: "Cell" }),
        cell("", { styleId: "Cell" }),
        cell("", { styleId: "Cell" }),
      ],
    ),
  );

  rows.push(row([cell("", { styleId: "Cell" })]));

  rows.push(
    row(
      [
        cell("Дата начала", { styleId: "Header" }),
        cell("Время", { styleId: "Header" }),
        cell("Преподаватели", { styleId: "Header" }),
        cell("Место обучения", { styleId: "Header" }),
        cell("Длительность", { styleId: "Header" }),
      ],
      { height: 22 },
    ),
  );

  placements.forEach((p) => {
    const teachers = p.teacherIds
      .map((id) => data.teachers.find((t) => t.id === id)?.name ?? "")
      .filter(Boolean)
      .join(", ");
    const location =
      data.locations.find((l) => l.id === p.locationId)?.name ?? "—";
    const startDate = parseISODate(p.startDate);
    const startWeekday = weekdayLabel(startDate);
    const time =
      p.customStartTime && p.customEndTime
        ? `${p.customStartTime} – ${p.customEndTime} (своё)`
        : p.timeSlotId
          ? (() => {
              const slot = data.settings.timeSlots.find(
                (s) => s.id === p.timeSlotId,
              );
              return slot
                ? `${slot.start} – ${slot.end}`
                : defaultRange;
            })()
          : defaultRange;
    rows.push(
      row([
        cell(`${p.startDate} (${startWeekday})`, { styleId: "CellCenter" }),
        cell(time, { styleId: "CellCenter" }),
        cell(teachers || "—", { styleId: "Cell" }),
        cell(location, { styleId: "Cell" }),
        cell(`${cycleDays} уч. дн.`, { styleId: "CellCenter" }),
      ]),
    );
  });

  rows.push(row([cell("", { styleId: "Cell" })]));

  rows.push(
    row(
      [
        cell("День цикла", { styleId: "Header" }),
        cell("Тема", { styleId: "Header" }),
        cell("ID", { styleId: "Header" }),
      ],
      { height: 22 },
    ),
  );

  data.topics.forEach((t, idx) => {
    rows.push(
      row([
        cell(idx + 1, { styleId: "CellCenter" }),
        cell(t.name, { styleId: "Cell" }),
        cell(t.id, { styleId: "CellCenter" }),
      ]),
    );
  });

  return buildSheet({
    name: group.name.slice(0, 31),
    rows,
    columnWidths: [180, 140, 280, 220, 120],
  });
}

function buildMasterCalendarSheet(data: AppData): string {
  const rows: string[] = [];

  rows.push(
    row(
      [
        cell("Сводный календарь занятий", { styleId: "Title" }),
        cell("", { styleId: "Title" }),
        cell("", { styleId: "Title" }),
      ],
      { height: 30 },
    ),
  );
  rows.push(row([cell("", { styleId: "Cell" })]));

  const groups = data.groups;
  if (groups.length === 0 || data.placements.length === 0) {
    rows.push(
      row([
        cell("Нет размещений для отображения", { styleId: "Cell" }),
        cell("", { styleId: "Cell" }),
        cell("", { styleId: "Cell" }),
      ]),
    );
    return buildSheet({
      name: "Сводный календарь",
      rows,
      columnWidths: [120, 280, 280],
    });
  }

  const startDate = parseISODate(data.settings.startDate);
  const placementsByStart = new Map<string, Placement[]>();
  data.placements.forEach((p) => {
    const list = placementsByStart.get(p.startDate) ?? [];
    list.push(p);
    placementsByStart.set(p.startDate, list);
  });

  const endDate = new Date(startDate);
  endDate.setMonth(endDate.getMonth() + 6);
  const cursor = new Date(startDate);

  const headerCells: string[] = [
    cell("Дата", { styleId: "Header" }),
    cell("День недели", { styleId: "Header" }),
    cell("День цикла", { styleId: "Header" }),
    ...groups.map((g, idx) =>
      cell(g.name, { styleId: `GroupHeader${idx % COLORS.length}` }),
    ),
  ];
  rows.push(row(headerCells, { height: 24 }));

  while (cursor <= endDate) {
    const iso = `${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}-${String(cursor.getDate()).padStart(2, "0")}`;
    const sunday = isSunday(cursor);

    if (sunday) {
      rows.push(
        row([
          cell(iso, { styleId: "Sunday" }),
          cell(weekdayLabel(cursor), { styleId: "Sunday" }),
          cell("выходной", { styleId: "Sunday" }),
          ...groups.map(() => cell("", { styleId: "Sunday" })),
        ]),
      );
    } else {
      const entries = getEntriesForDate(data, iso);
      const cycleDay = entries.length > 0 ? entries[0].cycleDayNumber : "";
      const groupCells: string[] = groups.map((g) => {
        const entry = entries.find((e) => e.groupId === g.id);
        if (!entry) return cell("", { styleId: "Cell" });
        const topic = entry.topicName ?? "—";
        const teachers = entry.teacherNames.join(", ");
        const location = entry.locationName ?? "";
        const time =
          entry.startTime && entry.endTime
            ? entry.timeSource === "custom"
              ? `${entry.startTime} – ${entry.endTime} (своё)`
              : entry.timeSource === "slot"
                ? `${entry.startTime} – ${entry.endTime} (слот)`
                : `${entry.startTime} – ${entry.endTime}`
            : "";
        const text = [
          topic,
          time && `Время: ${time}`,
          teachers && `Преп.: ${teachers}`,
          location && `Место: ${location}`,
        ]
          .filter(Boolean)
          .join("\n");
        return cell(text, { styleId: "Cell" });
      });

      rows.push(
        row([
          cell(iso, { styleId: "CellCenter" }),
          cell(weekdayLabel(cursor), { styleId: "Cell" }),
          cell(String(cycleDay), { styleId: "CellCenter" }),
          ...groupCells,
        ]),
      );
    }

    cursor.setDate(cursor.getDate() + 1);
  }

  return buildSheet({
    name: "Сводный календарь",
    rows,
    columnWidths: [120, 140, 80, ...groups.map(() => 220)],
  });
}

function buildMonthlyCalendarSheet(
  data: AppData,
  year: number,
  month: number,
): string {
  const rows: string[] = [];
  const monthName = [
    "Январь",
    "Февраль",
    "Март",
    "Апрель",
    "Май",
    "Июнь",
    "Июль",
    "Август",
    "Сентябрь",
    "Октябрь",
    "Ноябрь",
    "Декабрь",
  ][month];

  rows.push(
    row(
      [
        cell(`${monthName} ${year}`, { styleId: "Title" }),
        cell("", { styleId: "Title" }),
        cell("", { styleId: "Title" }),
        cell("", { styleId: "Title" }),
      ],
      { height: 30 },
    ),
  );

  rows.push(
    row([
      cell("Пн", { styleId: "Header" }),
      cell("Вт", { styleId: "Header" }),
      cell("Ср", { styleId: "Header" }),
      cell("Чт", { styleId: "Header" }),
      cell("Пт", { styleId: "Header" }),
      cell("Сб", { styleId: "Header" }),
      cell("Вс", { styleId: "Header" }),
    ]),
  );

  const weeks = buildMonthCalendar(year, month);
  weeks.forEach((week) => {
    const cells = week.days.map((d) => {
      if (!d) return cell("", { styleId: "Cell" });
      if (d.isSunday) {
        return cell(`${d.day}\nвыходной`, { styleId: "Sunday" });
      }
      const entries = getEntriesForDate(data, d.iso);
      if (entries.length === 0) {
        return cell(String(d.day), { styleId: "CellCenter" });
      }
      const lines = entries.map((e) => {
        const time =
          e.startTime && e.endTime ? `${e.startTime}–${e.endTime} ` : "";
        return `${e.groupName}: ${time}${e.topicName ?? "—"} | ${e.teacherNames.join(", ") || "—"} | ${e.locationName ?? "—"}`;
      });
      return cell(`${d.day}\n${lines.join("\n")}`, { styleId: "Cell" });
    });
    rows.push(row(cells));
  });

  return buildSheet({
    name: `${monthName} ${year}`.slice(0, 31),
    rows,
    columnWidths: [220, 220, 220, 220, 220, 220, 140],
  });
}

export function exportToExcel(data: AppData): void {
  const sheets: string[] = [];

  sheets.push(buildSettingsSheet(data));
  sheets.push(buildOverviewSheet(data));

  if (data.groups.length > 0 && data.placements.length > 0) {
    sheets.push(buildMasterCalendarSheet(data));
  }

  data.groups.forEach((g, idx) => {
    sheets.push(buildGroupScheduleSheet(g, data, idx));
  });

  const startDate = parseISODate(data.settings.startDate);
  for (let i = 0; i < 3; i += 1) {
    const d = new Date(startDate);
    d.setMonth(d.getMonth() + i);
    sheets.push(buildMonthlyCalendarSheet(data, d.getFullYear(), d.getMonth()));
  }

  const xml = `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
  ${buildStyles()}
  ${sheets.join("\n")}
</Workbook>`;

  const blob = new Blob([xml], {
    type: "application/vnd.ms-excel",
  });
  triggerDownload(blob, `raspisanie_${data.settings.startDate}.xls`);
}

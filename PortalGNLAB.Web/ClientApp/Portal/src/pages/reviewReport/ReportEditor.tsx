import type { ReportDoc } from "./types";

type Field = {
  label: string;
  fields?: Record<string, Field>;
  item?: Field;
  multiline?: boolean;
};
const text = (label: string): Field => ({ label });
const lines = (label: string): Field => ({ label, item: text("Line") });
const schema: Record<string, Field> = {
  title: text("Title"),
  productModel: text("Product model"),
  reportDate: text("Report date"),
  purpose: { label: "Purpose", multiline: true },
  background: { label: "Background", multiline: true },
  content: lines("Content"),
  resultTitle: text("Result title"),
  changedFactors: {
    label: "Changed factors",
    item: {
      label: "Factor",
      fields: {
        name: text("Name"),
        baseline: text("Baseline"),
        changed: text("Changed"),
      },
    },
  },
  conditions: {
    label: "Conditions",
    item: {
      label: "Condition",
      fields: {
        id: text("ID"),
        label: text("Label"),
        role: text("Role (control / test)"),
        date: text("Date"),
        model: text("Model"),
        details: text("Details"),
        table: text("Table"),
      },
    },
  },
  resultTables: {
    label: "Result tables",
    item: {
      label: "Table",
      fields: {
        caption: text("Caption"),
        columns: lines("Columns"),
        rows: { label: "Rows", item: lines("Cells") },
        source: text("Source"),
      },
    },
  },
  metrics: {
    label: "Metrics",
    item: {
      label: "Metric",
      fields: {
        id: text("ID"),
        name: text("Name"),
        unit: text("Unit"),
        kind: text("Kind"),
        rateOf: text("Rate of"),
        values: {
          label: "Values",
          item: {
            label: "Value",
            fields: { armId: text("Condition ID"), value: text("Value") },
          },
        },
      },
    },
  },
  findings: lines("Findings"),
  conclusion: { label: "Conclusion", multiline: true },
  limitations: lines("Limitations"),
  nextActions: lines("Next actions"),
  tags: lines("Tags"),
  images: {
    label: "Picture placements",
    item: {
      label: "Picture",
      fields: {
        no: text("Image number"),
        caption: text("Caption"),
        group: text("Group"),
        section: text("Section"),
      },
    },
  },
};

function emptyValue(field: Field): unknown {
  if (field.item) return [];
  if (field.fields)
    return Object.fromEntries(
      Object.entries(field.fields).map(([key, child]) => [
        key,
        key === "no" ? 1 : emptyValue(child),
      ]),
    );
  return "";
}

function FormField({
  field,
  value,
  onChange,
}: {
  field: Field;
  value: unknown;
  onChange: (value: unknown) => void;
}) {
  if (field.item) {
    const items = (value ?? []) as unknown[];
    return (
      <fieldset className="report-editor-group">
        <legend>{field.label}</legend>
        {items.map((item, index) => (
          <div key={index}>
            <FormField
              field={field.item!}
              value={item}
              onChange={(next) =>
                onChange(items.map((v, i) => (i === index ? next : v)))
              }
            />
            <button
              className="btn sm danger"
              onClick={() => onChange(items.filter((_, i) => i !== index))}
            >
              Remove {field.item!.label.toLowerCase()} {index + 1}
            </button>
          </div>
        ))}
        <button
          className="btn sm"
          onClick={() => onChange([...items, emptyValue(field.item!)])}
        >
          Add {field.item.label.toLowerCase()}
        </button>
      </fieldset>
    );
  }
  if (field.fields) {
    const record = (value ?? {}) as Record<string, unknown>;
    return (
      <fieldset className="report-editor-group">
        <legend>{field.label}</legend>
        {Object.entries(field.fields).map(([key, child]) => (
          <FormField
            key={key}
            field={child}
            value={record[key]}
            onChange={(next) => onChange({ ...record, [key]: next })}
          />
        ))}
      </fieldset>
    );
  }
  return (
    <label className="tool-field">
      {field.label}
      {field.multiline ? (
        <textarea
          className="inp tool-textarea"
          value={String(value ?? "")}
          onChange={(e) => onChange(e.target.value)}
        />
      ) : (
        <input
          className="inp"
          type={typeof value === "number" ? "number" : "text"}
          value={String(value ?? "")}
          onChange={(e) =>
            onChange(
              typeof value === "number"
                ? Number(e.target.value)
                : e.target.value,
            )
          }
        />
      )}
    </label>
  );
}

export default function ReportEditor({
  doc,
  onChange,
}: {
  doc: ReportDoc;
  onChange: (doc: ReportDoc) => void;
}) {
  return (
    <details className="report-editor">
      <summary>Edit report details</summary>
      <p className="muted">
        Edit fields, conditions and measured values. Save to keep changes.
      </p>
      {Object.entries(schema).map(([key, field]) => (
        <FormField
          key={key}
          field={field}
          value={doc[key as keyof ReportDoc]}
          onChange={(value) => onChange({ ...doc, [key]: value })}
        />
      ))}
    </details>
  );
}

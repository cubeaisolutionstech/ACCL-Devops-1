
import React from "react";

interface ToggleButtonGroupProps {
  options: { label: string; value: string }[];
  value: string;
  onChange: (value: string) => void;
}

const ToggleButtonGroup: React.FC<ToggleButtonGroupProps> = ({
  options,
  value,
  onChange,
}) => {
  return (
    <div className="flex w-full max-w-lg rounded-lg border border-blue-200 overflow-hidden">
      {options.map((option, idx) => {
        const selected = value === option.value;
        return (
          <button
            key={option.value}
            onClick={() => onChange(option.value)}
            className={[
              "flex-1 px-6 py-3 font-semibold text-base focus:outline-none transition-colors duration-150",
              selected
                ? "bg-blue-600 text-white"
                : "bg-white text-blue-800 border-transparent",
              idx === 0
                ? "rounded-l-lg"
                : "",
              idx === options.length - 1
                ? "rounded-r-lg"
                : "",
              "border-0",
            ].join(" ")}
            style={{
              borderRight:
                idx < options.length - 1 ? "1px solid #e5e7eb" : undefined,
            }}
            type="button"
          >
            {option.label}
          </button>
        );
      })}
    </div>
  );
};

export default ToggleButtonGroup;

/* 
// Usage Example:
import React, { useState } from "react";
import ToggleButtonGroup from "@/components/ToggleButtonGroup";

export default function Example() {
  const [selected, setSelected] = useState("executive");
  return (
    <ToggleButtonGroup
      options={[
        { label: "Executive Creation", value: "executive" },
        { label: "Customer Code Management", value: "customer" },
      ]}
      value={selected}
      onChange={setSelected}
    />
  );
}
*/

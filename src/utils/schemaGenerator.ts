import { z } from 'zod';
import { ColumnConfig, ListConfig } from '../types';

/**
 * Enterprise Zod Schema Factory for Dynamic Forms
 *
 * Dynamically constructs a Zod validation schema based on SharePoint Column Configurations.
 * Ensures strict runtime type checking and field requirement validation without
 * hardcoding form schemas across different SharePoint list definitions.
 *
 * @param columnsOrConfig - Array of ColumnConfig or a ListConfig contract
 * @returns ZodObject dynamic validation schema
 */
export function generateZodSchema(
  columnsOrConfig: ColumnConfig[] | ListConfig
): z.ZodObject<any> {
  const columns = Array.isArray(columnsOrConfig)
    ? columnsOrConfig
    : columnsOrConfig.columns || [];

  const shape: Record<string, z.ZodTypeAny> = {};

  columns.forEach((col) => {
    // Only construct validation fields for editable or user-supplied columns
    const isRequired = Boolean(col.required);
    const label = col.label || col.key;

    switch (col.type) {
      case 'number': {
        const numSchema = z.coerce.number();

        if (isRequired) {
          shape[col.key] = numSchema;
        } else {
          shape[col.key] = z
            .union([z.number(), z.string().length(0), z.null(), z.undefined()])
            .transform((val) => (val === '' || val === null || val === undefined ? undefined : Number(val)))
            .optional();
        }
        break;
      }

      case 'boolean': {
        shape[col.key] = z.boolean().optional().default(Boolean(col.defaultValue ?? false));
        break;
      }

      case 'choice': {
        if (isRequired) {
          shape[col.key] = z
            .string()
            .min(1, `${label} is required`);
        } else {
          shape[col.key] = z.string().optional().or(z.literal(''));
        }
        break;
      }

      case 'multichoice': {
        if (isRequired) {
          shape[col.key] = z
            .array(z.string())
            .min(1, `At least one selection for ${label} is required`);
        } else {
          shape[col.key] = z.array(z.string()).optional().default([]);
        }
        break;
      }

      case 'date':
      case 'datetime': {
        if (isRequired) {
          shape[col.key] = z
            .string()
            .min(1, `${label} is required`);
        } else {
          shape[col.key] = z.string().optional().or(z.literal(''));
        }
        break;
      }

      case 'text':
      default: {
        if (isRequired) {
          shape[col.key] = z
            .string()
            .min(1, `${label} is required`);
        } else {
          shape[col.key] = z.string().optional().or(z.literal(''));
        }
        break;
      }
    }
  });

  return z.object(shape);
}

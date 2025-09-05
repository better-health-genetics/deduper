
export interface AddRecordFormData {
  firstName: string;
  lastName: string;
  dob: string;
}

export interface AddRecordResponse {
  success: boolean;
  message: string;
  duplicatesFound?: number;
}

export interface DuplicateHealthData {
  percentage: number;
  totalChecked: number;
  totalDuplicates: number;
}

export enum MessageType {
  Success,
  Error,
  Warning,
  None,
}


import type { AddRecordFormData, AddRecordResponse, DuplicateHealthData } from '../types';

// This file mocks the `google.script.run` functionality for local development.
// In a real Google Apps Script web app, you would call the server-side functions directly.

const MOCK_LATENCY = 800; // ms

/**
 * Mocks the server-side `addRecordAndCheckDuplicates` function.
 */
export const addRecordAndCheckDuplicates = (formData: AddRecordFormData): Promise<AddRecordResponse> => {
  console.log('Simulating addRecordAndCheckDuplicates with:', formData);

  return new Promise((resolve, reject) => {
    setTimeout(() => {
      if (!formData.firstName || !formData.lastName || !formData.dob) {
        resolve({
          success: false,
          message: 'Error: All mandatory fields must be provided.',
        });
        return;
      }
      
      // Simulate a random outcome
      const isDuplicate = Math.random() > 0.6; // 40% chance of being a duplicate
      if (isDuplicate) {
        const numDupes = Math.floor(Math.random() * 4) + 1; // 1 to 4 duplicates
        resolve({
          success: false,
          message: `Record added to "Checker" tab. ${numDupes} duplicates found.`,
          duplicatesFound: numDupes,
        });
      } else {
        resolve({
          success: true,
          message: 'Record added to "Checker" tab. No duplicates found.',
        });
      }
    }, MOCK_LATENCY);
  });
};

/**
 * Mocks a new server-side function to get duplicate health statistics.
 * This function would need to be implemented in your Code.gs file.
 * It calculates the percentage of duplicates since the previous Monday.
 */
export const getDuplicateHealthData = (): Promise<DuplicateHealthData> => {
  console.log('Simulating getDuplicateHealthData...');
  
  return new Promise((resolve) => {
    setTimeout(() => {
      // Simulate some realistic data
      const totalChecked = Math.floor(Math.random() * 150) + 10; // 10 to 160
      const totalDuplicates = Math.floor(Math.random() * (totalChecked * 0.45)); // up to 45% duplicates
      const percentage = totalChecked > 0 ? Math.round((totalDuplicates / totalChecked) * 100) : 0;
      
      resolve({
        percentage,
        totalChecked,
        totalDuplicates,
      });
    }, MOCK_LATENCY);
  });
};

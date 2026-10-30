import React from 'react';
import { useAuth } from '../auth/AuthContext';

/** Shows its content to admins only (used for Drive / Google Sheet links). */
export const AdminOnly: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { isAdmin } = useAuth();
  return isAdmin ? <>{children}</> : null;
};

import { create } from "zustand";

export interface UserProfile {
  id: string;
  name: string;
  email: string;
  avatarUrl?: string;
}

interface AuthState {
  user: UserProfile | null;
  token: string | null;
  isAuthenticated: boolean;
  login: (profile: UserProfile, token: string) => void;
  logout: () => void;
  initializeFromStorage: () => void;
}

const STORAGE_KEY = "nlc-auth";

export const useAuthStore = create<AuthState>((set) => ({
  user: null,
  token: null,
  isAuthenticated: false,
  login: (profile, token) => {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({ profile, token, timestamp: Date.now() })
    );
    set({ user: profile, token, isAuthenticated: true });
  },
  logout: () => {
    localStorage.removeItem(STORAGE_KEY);
    set({ user: null, token: null, isAuthenticated: false });
  },
  initializeFromStorage: () => {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    try {
      const parsed = JSON.parse(raw) as {
        profile: UserProfile;
        token: string;
        timestamp: number;
      };
      const oneDay = 24 * 60 * 60 * 1000;
      if (Date.now() - parsed.timestamp < oneDay) {
        set({
          user: parsed.profile,
          token: parsed.token,
          isAuthenticated: true
        });
      } else {
        localStorage.removeItem(STORAGE_KEY);
      }
    } catch (error) {
      console.error("Failed to parse auth storage", error);
      localStorage.removeItem(STORAGE_KEY);
    }
  }
}));

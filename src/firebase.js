import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyAuzgZscO2klQVDX0fbA0jYyziMkox0Emg",
  authDomain: "accountbook-1c2d5.firebaseapp.com",
  projectId: "accountbook-1c2d5",
  storageBucket: "accountbook-1c2d5.firebasestorage.app",
  messagingSenderId: "841321241774",
  appId: "1:841321241774:web:c2c31f89c76802fec7a13f",
  measurementId: "G-RW543W6HVV"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
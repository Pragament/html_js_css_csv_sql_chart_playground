import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getAuth, signInWithPopup, GoogleAuthProvider, signOut, signInWithRedirect, getRedirectResult } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import { getFirestore, doc, setDoc, getDoc, collection, query, getDocs } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const firebaseConfig = {
    apiKey: "AIzaSyAp3Zj772z5kozR577L9UyN7CPJ6jdCUBg",
    authDomain: "coherent-flame-435810-q5.firebaseapp.com",
    projectId: "coherent-flame-435810-q5",
    storageBucket: "coherent-flame-435810-q5.firebasestorage.app",
    messagingSenderId: "914114495752",
    appId: "1:914114495752:web:17db8d870adb78033ad8bd",
    measurementId: "G-CZWSYGJBDV"
};

try {
    const app = initializeApp(firebaseConfig);
    const auth = getAuth(app);
    const db = getFirestore(app);
    const provider = new GoogleAuthProvider();
    provider.addScope('email');

    window.firebaseAuth = auth;
    window.firebaseDb = db;
    window.firebaseProvider = provider;
    window.signInWithPopup = signInWithPopup;
    window.signInWithRedirect = signInWithRedirect;
    window.getRedirectResult = getRedirectResult;
    window.signOut = signOut;
    window.doc = doc;
    window.setDoc = setDoc;
    window.getDoc = getDoc;
    window.collection = collection;
    window.query = query;
    window.getDocs = getDocs;

    console.log("✅ Firebase initialized");
} catch (error) {
    console.error("Firebase initialization failed:", error);
    alert("Failed to initialize Firebase: " + error.message);
}
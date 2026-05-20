const API_URL = "https://app-generacion-informes.onrender.com/datos";

export const obtenerDatos = async () => {
    try {

        const response = await fetch(API_URL);

        console.log("STATUS:", response.status);

        const data = await response.json();

        console.log("DATOS:", data);

        return Array.isArray(data) ? data : [];

    } catch (error) {

        console.error("ERROR:", error);

        return [];
    }
};
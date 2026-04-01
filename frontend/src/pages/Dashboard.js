import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { obtenerDatos } from "../services/api";
import tecnicos from "../data/tecnicos.json";

function Dashboard() {

    const navigate = useNavigate();
    const [datos, setDatos] = useState([]);
    const [seleccionados, setSeleccionados] =useState([]);
    const [cargando, setCargando] = useState(false);

    //se valida datos
    useEffect(() =>{
        const usuario = localStorage.getItem("usuario");

        if(!usuario){
            navigate("/", {replace: true});
        };
    }, [navigate]);

    //se traen datos de epicoollect
    useEffect(() =>{
        const cargarDatos = async () => {
            const res = await obtenerDatos();
            setDatos(res);
        };

        cargarDatos();
    }, []);

    // Cerrar Sesión
    const logout =() =>{
        localStorage.removeItem("usuario");
        navigate("/",{replace: true})
    };

    //caja de selección
    const handleSelection = (item) => {
        const id = item.ec5_uuid;

        if (!id) {
            console.error("ID inválido:", item);
            return;
        }

        setSeleccionados(prev => {
            if (prev.includes(id)) {
                return prev.filter(i => i !== id);
            } else {
                return [...prev, id];
            }
        });
    };

    //traer nombre de tecnico

    const traerIdTecnico = (id) =>{
        const tecnico = tecnicos.find(t => t.id === id);
        return tecnico ? tecnico.nombre : "No encontrado"
    }

    //boton descargar todos los archivos
    const descargarTodos = async () => {
        setCargando(true);

        try {
            const response = await fetch("http://127.0.0.1:5000/generar", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ ids: [] })
            });

            const blob = await response.blob();

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;

            const contentDisposition = response.headers.get("Content-Disposition");
            let nombreArchivo = "informes.zip";

            if (contentDisposition) {
                const match = contentDisposition.match(/filename="(.+)"/);
                if (match) {
                    nombreArchivo = match[1];
                }
            }

            a.download = nombreArchivo;
            a.click();

        } catch (error) {
            console.error(error);
        } finally {
            setCargando(false); // 🔥 clave
        }
    };

    //boton descargar seleccionados
    const descargarSeleccionados = async () => {
        setCargando(true);

        try {
            const response = await fetch("http://127.0.0.1:5000/generar", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ ids: seleccionados })
            });

            const blob = await response.blob();

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;

            const contentDisposition = response.headers.get("Content-Disposition");
            let nombreArchivo = "informes.zip";

            if (contentDisposition) {
                const match = contentDisposition.match(/filename="(.+)"/);
                if (match) {
                    nombreArchivo = match[1];
                }
            }

            a.download = nombreArchivo;
            a.click();

        } catch (error) {
            console.error(error);
        } finally {
            setCargando(false); // 🔥 clave
        }
    };

    

    return (
    <div>
        <h1>Dashboard</h1>
        <p>Bienvenido al sistema de informes UPS</p>

        <button onClick={logout}>Cerrar Sesión</button>

        <h2>Listado de Informes</h2>

        {datos.length === 0 ? (
            <p>Cargando datos...</p>
        ) : (
            <table border="1" cellPadding="10" style={{ marginTop: "20px", borderCollapse: "collapse" }}>
                <thead>
                    <tr>
                        <th>Selección</th>
                        <th>Título</th>
                        <th>Fecha</th>
                        <th>Ciudad</th>
                        <th>Departamento</th>
                        <th>Capacidad UPS</th>
                        <th>Técnico</th>
                        <th>Nombre Sede</th>
                        <th>Dirección</th>
                    </tr>
                </thead>

                <tbody>
                    {datos.map((item, index) => (
                        <tr key={index}>
                            <td><input 
                                type="checkbox" 
                                checked={seleccionados.includes(item.ec5_uuid)}
                                onChange={() => handleSelection(item)}
                            /></td>
                            <td>{item.title}</td>
                            <td>{item.created_at.split("T")[0]}</td>
                            <td>{item["7_CIUDAD"]}</td>
                            <td>{item["8_DEPARTAMENTO"]}</td>
                            <td>{item["66_CAPACIDAD_UPS_KVA"]+" KVA"}</td>
                            <td>{traerIdTecnico(item["11_CODIGO_TECNICO"])}</td>
                            <td>{item["5_NOMBRE_SEDE"]}</td>
                            <td>{item["6_DIRECCION"]}</td>
                        </tr>                        
                    ))}
                </tbody>
            </table>
        )}
        <button onClick={descargarTodos} disabled={cargando}>
            {cargando ? "Generando..." : "Descargar Todos"}
        </button>

        <button onClick={descargarSeleccionados} disabled={cargando}>
            {cargando ? "Generando..." : "Descargar Seleccionados"}
        </button>
        {cargando && (
            <p style={{ color: "blue", fontWeight: "bold", marginTop: "10px" }}>
                ⏳ Generando y descargando informes, por favor espera...
            </p>
        )}
    </div>
    );
};

export default Dashboard;
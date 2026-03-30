
import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { obtenerDatos } from "../services/api";
import tecnicos from "../data/tecnicos.json";

function Dashboard() {

    const navigate = useNavigate();
    const [datos, setDatos] = useState([]);
    const [seleccionados, setSeleccionados] =useState([]);

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
    const handleSelection =(item) =>{
        const id = item.ec5_uuid
        const existe = seleccionados.includes(id);

        if(existe){
            setSeleccionados(seleccionados.filter(i => i!== id));
        } else{
            setSeleccionados([...seleccionados, id]);
        }
    };

    //traer nombre de tecnico

    const traerIdTecnico = (id) =>{
        const tecnico = tecnicos.find(t => t.id === id);
        return tecnico ? tecnico.nombre : "No encontrado"
    }

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
                            <td><input type="checkbox" onChange={()=>handleSelection(item)}/></td>
                            <td>{item.title}</td>
                            <td>{item.created_at.split("T")[0]}</td>
                            <td>{item["7_CIUDAD"]}</td>
                            <td>{item["8_DEPARTAMENTO"]}</td>
                            <td>{item["63_CAPACIDAD_UPS_KVA"]+" KVA"}</td>
                            <td>{traerIdTecnico(item["11_CODIGO_TECNICO"])}</td>
                            <td>{item["5_NOMBRE_SEDE"]}</td>
                            <td>{item["6_DIRECCION"]}</td>
                        </tr>
                    ))}
                </tbody>
            </table>
        )}
        <button onClick={()=> console.log(seleccionados)}>Ver Seleccionados</button>
    </div>
    );
};

export default Dashboard;
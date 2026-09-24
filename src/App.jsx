import React, { useState, useMemo } from 'react';
import { 
  Building2, 
  FileText, 
  Upload, 
  Users, 
  ShieldAlert, 
  AlertTriangle, 
  CheckCircle2, 
  XCircle, 
  Search, 
  ExternalLink, 
  LogOut, 
  UserPlus, 
  Plus, 
  Download,
  FolderOpen,
  Lock,
  ChevronRight,
  ShieldCheck,
  FileSpreadsheet,
  KeyRound
} from 'lucide-react';
import * as XLSX from 'xlsx';

// Usuarios de prueba iniciales
const INITIAL_USERS = [
  { id: '1', email: 'julio.benages@neural.es', code: 'N3UR4L2026X', name: 'Julio Benages', role: 'superadmin', assignedCentres: ['ALL'] },
  { id: '2', email: 'tecnico.madrid@neural.es', code: 'M4DR1D2026', name: 'Técnico Madrid', role: 'gestor', assignedCentres: ['c1'] },
  { id: '3', email: 'tecnico.valencia@neural.es', code: 'V4L3NC1426', name: 'Técnico Valencia', role: 'gestor', assignedCentres: ['c2'] },
];

const INITIAL_GENERAL_DOCS = [
  { id: 'gd1', title: 'Procedimiento General de Evacuación v2', category: 'Procedimientos', link: 'https://sharepoint.com/doc1', date: '2026-01-15' },
  { id: 'gd2', title: 'Protocolo de Actuación Accidentes Laborales', category: 'Protocolos', link: 'https://sharepoint.com/doc2', date: '2026-02-01' },
  { id: 'gd3', title: 'Plantilla de Inspección de Equipos de Protección', category: 'Plantillas', link: 'https://sharepoint.com/doc3', date: '2026-02-10' },
];

const INITIAL_CENTRES = [
  {
    id: 'c1',
    name: 'Centro Neural Madrid - Castellana',
    zone: 'Madrid Norte',
    users: ['julio.benages@neural.es', 'tecnico.madrid@neural.es'],
    docs: {
      evaluacion_riesgos: { title: 'Evaluación de Riesgos Laborales 2026', link: 'https://sharepoint.com/eval-madrid', status: 'presente' },
      informacion_riesgos: { title: 'Ficha Informativa Puestos de Trabajo', link: 'https://sharepoint.com/info-madrid', status: 'presente' },
      medidas_emergencia: { title: '', link: '', status: 'pendiente' }
    }
  },
  {
    id: 'c2',
    name: 'Centro Neural Valencia - Mestalla',
    zone: 'Comunidad Valenciana',
    users: ['julio.benages@neural.es', 'tecnico.valencia@neural.es'],
    docs: {
      evaluacion_riesgos: { title: 'Evaluación de Riesgos 2025/2026', link: 'https://sharepoint.com/eval-valencia', status: 'presente' },
      informacion_riesgos: { title: '', link: '', status: 'pendiente' },
      medidas_emergencia: { title: 'Plan de Emergencia y Evacuación', link: 'https://sharepoint.com/emerg-valencia', status: 'presente' }
    }
  }
];

export default function App() {
  // Autenticación
  const [currentUser, setCurrentUser] = useState(null);
  const [loginEmail, setLoginEmail] = useState('');
  const [loginCode, setLoginCode] = useState('');
  const [loginError, setLoginError] = useState('');

  // Estado global
  const [users, setUsers] = useState(INITIAL_USERS);
  const [centres, setCentres] = useState(INITIAL_CENTRES);
  const [generalDocs, setGeneralDocs] = useState(INITIAL_GENERAL_DOCS);

  // Navegación
  const [activeTab, setActiveTab] = useState('centres');
  const [selectedCentre, setSelectedCentre] = useState(null);
  const [searchTerm, setSearchTerm] = useState('');

  // Modal enlace SharePoint
  const [editDocModal, setEditDocModal] = useState({ open: false, centreId: null, categoryKey: null, title: '', link: '' });

  // Manejo de Login
  const handleLogin = (e) => {
    e.preventDefault();
    setLoginError('');
    const user = users.find(
      u => u.email.toLowerCase().trim() === loginEmail.toLowerCase().trim() && 
           u.code.trim() === loginCode.trim()
    );

    if (user) {
      setCurrentUser(user);
    } else {
      setLoginError('Correo o código alfanumérico incorrectos.');
    }
  };

  const handleLogout = () => {
    setCurrentUser(null);
    setLoginEmail('');
    setLoginCode('');
  };

  const accessibleCentres = useMemo(() => {
    if (!currentUser) return [];
    if (currentUser.role === 'superadmin') return centres;
    return centres.filter(c => c.users.includes(currentUser.email));
  }, [currentUser, centres]);

  // Carga Masiva Excel
  const handleFileUploadCentres = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

      const rows = data.slice(1);
      const newCentres = [...centres];
      const newUsers = [...users];

      rows.forEach((row, idx) => {
        if (!row[0]) return;

        const centreName = row[0].toString().trim();
        const zone = row[1] ? row[1].toString().trim() : 'General';
        const assignedUserEmails = [];

        // Leer pares (C, D), (E, F), (G, H), (I, J), (K, L)
        const userCols = [
          { emailIdx: 2, codeIdx: 3 },
          { emailIdx: 4, codeIdx: 5 },
          { emailIdx: 6, codeIdx: 7 },
          { emailIdx: 8, codeIdx: 9 },
          { emailIdx: 10, codeIdx: 11 }
        ];

        userCols.forEach(({ emailIdx, codeIdx }) => {
          if (row[emailIdx] && row[codeIdx]) {
            const email = row[emailIdx].toString().trim();
            const code = row[codeIdx].toString().trim();
            assignedUserEmails.push(email);

            // Crear usuario si no existe
            const exists = newUsers.some(u => u.email.toLowerCase() === email.toLowerCase());
            if (!exists) {
              newUsers.push({
                id: `u_${Date.now()}_${Math.random()}`,
                email: email,
                code: code,
                name: email.split('@')[0],
                role: 'gestor',
                assignedCentres: []
              });
            }
          }
        });

        const existingIndex = newCentres.findIndex(c => c.name.toLowerCase() === centreName.toLowerCase());
        const centreObj = {
          id: existingIndex >= 0 ? newCentres[existingIndex].id : `c_${Date.now()}_${idx}`,
          name: centreName,
          zone: zone,
          users: assignedUserEmails,
          docs: existingIndex >= 0 ? newCentres[existingIndex].docs : {
            evaluacion_riesgos: { title: '', link: '', status: 'pendiente' },
            informacion_riesgos: { title: '', link: '', status: 'pendiente' },
            medidas_emergencia: { title: '', link: '', status: 'pendiente' }
          }
        };

        if (existingIndex >= 0) {
          newCentres[existingIndex] = centreObj;
        } else {
          newCentres.push(centreObj);
        }
      });

      setCentres(newCentres);
      setUsers(newUsers);
      alert('¡Centros y usuarios importados correctamente!');
    };
    reader.readAsBinaryString(file);
  };

  const saveSharepointLink = () => {
    if (!editDocModal.centreId || !editDocModal.categoryKey) return;

    setCentres(prevCentres => prevCentres.map(c => {
      if (c.id === editDocModal.centreId) {
        const isPresent = editDocModal.link.trim() !== '';
        return {
          ...c,
          docs: {
            ...c.docs,
            [editDocModal.categoryKey]: {
              title: editDocModal.title || (isPresent ? 'Documento Vinculado' : ''),
              link: editDocModal.link,
              status: isPresent ? 'presente' : 'pendiente'
            }
          }
        };
      }
      return c;
    }));

    setEditDocModal({ open: false, centreId: null, categoryKey: null, title: '', link: '' });
  };

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4 font-sans">
        <div className="bg-white rounded-xl shadow-2xl p-8 max-w-md w-full space-y-6">
          <div className="text-center space-y-2">
            <div className="bg-blue-600 text-white w-12 h-12 rounded-lg flex items-center justify-center mx-auto shadow-lg">
              <ShieldCheck className="w-8 h-8" />
            </div>
            <h1 className="text-2xl font-bold text-slate-800">Plataforma PRL</h1>
            <p className="text-sm text-slate-500">Gestión Documental y Prevención de Riesgos</p>
          </div>

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Correo Electrónico</label>
              <input 
                type="email" 
                required
                className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                placeholder="ejemplo@neural.es"
                value={loginEmail}
                onChange={(e) => setLoginEmail(e.target.value)}
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Código Alfanumérico (Contraseña)</label>
              <input 
                type="password" 
                required
                className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                placeholder="Tu código de acceso"
                value={loginCode}
                onChange={(e) => setLoginCode(e.target.value)}
              />
            </div>

            {loginError && (
              <div className="p-3 bg-red-50 border border-red-200 text-red-600 text-xs rounded-lg flex items-center gap-2">
                <AlertTriangle className="w-4 h-4 shrink-0" />
                <span>{loginError}</span>
              </div>
            )}

            <button 
              type="submit" 
              className="w-full py-3 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition shadow-md"
            >
              Iniciar Sesión
            </button>
          </form>

          <div className="text-xs text-center text-slate-400 border-t pt-4">
            Acceso restringido a personal autorizado por la empresa.
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col font-sans">
      <header className="bg-slate-800 text-white shadow-md">
        <div className="max-w-7xl mx-auto px-4 py-3 flex justify-between items-center">
          <div className="flex items-center space-x-3">
            <ShieldCheck className="w-7 h-7 text-blue-400" />
            <div>
              <h1 className="font-bold text-lg leading-tight">Neural PRL</h1>
              <span className="text-xs text-slate-400">Panel de Control de Prevención</span>
            </div>
          </div>

          <div className="flex items-center space-x-4">
            <div className="text-right hidden sm:block">
              <p className="text-sm font-medium">{currentUser.name}</p>
              <span className="text-xs bg-blue-900 text-blue-200 px-2 py-0.5 rounded capitalize">
                {currentUser.role}
              </span>
            </div>
            <button 
              onClick={handleLogout} 
              className="p-2 text-slate-300 hover:text-white hover:bg-slate-700 rounded-lg transition"
              title="Cerrar Sesión"
            >
              <LogOut className="w-5 h-5" />
            </button>
          </div>
        </div>
      </header>

      <nav className="bg-white border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-4 flex space-x-8">
          <button 
            onClick={() => { setActiveTab('centres'); setSelectedCentre(null); }}
            className={`py-4 px-2 font-medium text-sm border-b-2 flex items-center space-x-2 ${activeTab === 'centres' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
          >
            <Building2 className="w-4 h-4" />
            <span>Mis Centros de Trabajo</span>
          </button>

          <button 
            onClick={() => setActiveTab('general')}
            className={`py-4 px-2 font-medium text-sm border-b-2 flex items-center space-x-2 ${activeTab === 'general' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
          >
            <FolderOpen className="w-4 h-4" />
            <span>Documentación General</span>
          </button>

          {currentUser.role === 'superadmin' && (
            <>
              <button 
                onClick={() => setActiveTab('excel')}
                className={`py-4 px-2 font-medium text-sm border-b-2 flex items-center space-x-2 ${activeTab === 'excel' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
              >
                <FileSpreadsheet className="w-4 h-4" />
                <span>Carga Masiva (Excel)</span>
              </button>

              <button 
                onClick={() => setActiveTab('users')}
                className={`py-4 px-2 font-medium text-sm border-b-2 flex items-center space-x-2 ${activeTab === 'users' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
              >
                <Users className="w-4 h-4" />
                <span>Gestión de Perfiles</span>
              </button>
            </>
          )}
        </div>
      </nav>

      <main className="max-w-7xl mx-auto px-4 py-8 flex-1 w-full">
        {activeTab === 'centres' && !selectedCentre && (
          <div className="space-y-6">
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div>
                <h2 className="text-xl font-bold text-slate-800">Centros Asignados</h2>
                <p className="text-sm text-slate-500">Selecciona un centro para revisar o enlazar la documentación.</p>
              </div>

              <div className="relative w-full sm:w-64">
                <Search className="w-4 h-4 absolute left-3 top-3 text-slate-400" />
                <input 
                  type="text"
                  placeholder="Buscar centro..."
                  className="w-full pl-9 pr-4 py-2 border border-slate-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              {accessibleCentres
                .filter(c => c.name.toLowerCase().includes(searchTerm.toLowerCase()))
                .map((centre) => {
                  const pendingCount = Object.values(centre.docs).filter(d => d.status === 'pendiente').length;
                  return (
                    <div 
                      key={centre.id}
                      onClick={() => setSelectedCentre(centre)}
                      className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm hover:shadow-md transition cursor-pointer flex flex-col justify-between"
                    >
                      <div className="space-y-3">
                        <div className="flex justify-between items-start">
                          <span className="text-xs font-semibold px-2.5 py-1 bg-slate-100 text-slate-600 rounded">
                            {centre.zone}
                          </span>
                          {pendingCount === 0 ? (
                            <span className="text-xs font-semibold px-2.5 py-1 bg-emerald-100 text-emerald-700 rounded-full flex items-center gap-1">
                              <CheckCircle2 className="w-3 h-3" /> OK
                            </span>
                          ) : (
                            <span className="text-xs font-semibold px-2.5 py-1 bg-rose-100 text-rose-700 rounded-full flex items-center gap-1">
                              <AlertTriangle className="w-3 h-3" /> {pendingCount} Pendiente
                            </span>
                          )}
                        </div>

                        <h3 className="text-lg font-bold text-slate-800">
                          {centre.name}
                        </h3>
                      </div>

                      <div className="mt-6 pt-4 border-t border-slate-100 flex justify-between items-center text-sm font-medium text-blue-600">
                        <span>Ver documentación</span>
                        <ChevronRight className="w-4 h-4" />
                      </div>
                    </div>
                  );
                })}
            </div>
          </div>
        )}

        {activeTab === 'centres' && selectedCentre && (
          <div className="space-y-6">
            <button 
              onClick={() => setSelectedCentre(null)}
              className="text-sm font-medium text-slate-500 hover:text-slate-800 flex items-center gap-1"
            >
              ← Volver a centros
            </button>

            <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-2">
              <span className="text-xs font-semibold px-2.5 py-1 bg-slate-100 text-slate-600 rounded">
                {selectedCentre.zone}
              </span>
              <h2 className="text-2xl font-bold text-slate-800">{selectedCentre.name}</h2>
              <p className="text-xs text-slate-400">
                Usuarios con acceso: {selectedCentre.users.join(', ')}
              </p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              {[
                { key: 'evaluacion_riesgos', label: 'Evaluación de Riesgos' },
                { key: 'informacion_riesgos', label: 'Información de Riesgos' },
                { key: 'medidas_emergencia', label: 'Medidas de Emergencia' }
              ].map(({ key, label }) => {
                const doc = selectedCentre.docs[key];
                const isPresent = doc && doc.status === 'presente';

                return (
                  <div key={key} className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex flex-col justify-between space-y-4">
                    <div>
                      <div className="flex justify-between items-center mb-3">
                        <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">Categoría</span>
                        {isPresent ? (
                          <span className="text-xs font-bold px-2.5 py-1 bg-emerald-100 text-emerald-700 rounded-full flex items-center gap-1">
                            <CheckCircle2 className="w-3 h-3" /> Presente / OK
                          </span>
                        ) : (
                          <span className="text-xs font-bold px-2.5 py-1 bg-rose-100 text-rose-700 rounded-full flex items-center gap-1">
                            <XCircle className="w-3 h-3" /> Pendiente
                          </span>
                        )}
                      </div>

                      <h3 className="text-base font-bold text-slate-800 mb-2">{label}</h3>

                      {isPresent ? (
                        <p className="text-sm text-slate-600 bg-slate-50 p-3 rounded border border-slate-100">
                          {doc.title || 'Documento disponible en SharePoint'}
                        </p>
                      ) : (
                        <p className="text-xs text-rose-500 italic bg-rose-50 p-3 rounded border border-rose-100">
                          Aún no se ha vinculado ningún documento para esta categoría en SharePoint.
                        </p>
                      )}
                    </div>

                    <div className="space-y-2 pt-2 border-t border-slate-100">
                      {isPresent && (
                        <a 
                          href={doc.link} 
                          target="_blank" 
                          rel="noreferrer"
                          className="w-full py-2 px-3 bg-blue-50 text-blue-600 font-medium text-xs rounded-lg hover:bg-blue-100 transition flex items-center justify-center gap-2"
                        >
                          <ExternalLink className="w-3.5 h-3.5" />
                          Abrir en SharePoint
                        </a>
                      )}

                      {currentUser.role === 'superadmin' && (
                        <button 
                          onClick={() => setEditDocModal({
                            open: true,
                            centreId: selectedCentre.id,
                            categoryKey: key,
                            title: doc?.title || '',
                            link: doc?.link || ''
                          })}
                          className="w-full py-2 px-3 bg-slate-100 text-slate-700 font-medium text-xs rounded-lg hover:bg-slate-200 transition"
                        >
                          {isPresent ? 'Editar Enlace SharePoint' : '+ Enlazar Documento SharePoint'}
                        </button>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {activeTab === 'general' && (
          <div className="space-y-6">
            <div>
              <h2 className="text-xl font-bold text-slate-800">Documentación General de PRL</h2>
              <p className="text-sm text-slate-500">Procedimientos, normas, protocolos y plantillas comunes a todos los centros.</p>
            </div>

            <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <table className="w-full text-left text-sm text-slate-600">
                <thead className="bg-slate-50 text-xs font-semibold text-slate-500 uppercase tracking-wider border-b">
                  <tr>
                    <th className="px-6 py-3">Documento</th>
                    <th className="px-6 py-3">Categoría</th>
                    <th className="px-6 py-3">Fecha</th>
                    <th className="px-6 py-3 text-right">Acceso</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {generalDocs.map((doc) => (
                    <tr key={doc.id} className="hover:bg-slate-50">
                      <td className="px-6 py-4 font-medium text-slate-800">{doc.title}</td>
                      <td className="px-6 py-4">
                        <span className="px-2 py-1 bg-slate-100 text-slate-600 text-xs rounded">
                          {doc.category}
                        </span>
                      </td>
                      <td className="px-6 py-4 text-slate-400 text-xs">{doc.date}</td>
                      <td className="px-6 py-4 text-right">
                        <a 
                          href={doc.link} 
                          target="_blank" 
                          rel="noreferrer"
                          className="text-blue-600 hover:text-blue-800 text-xs font-semibold inline-flex items-center gap-1"
                        >
                          Ver archivo <ExternalLink className="w-3 h-3" />
                        </a>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {activeTab === 'excel' && currentUser.role === 'superadmin' && (
          <div className="max-w-2xl mx-auto bg-white p-8 rounded-xl border border-slate-200 shadow-sm space-y-6">
            <div>
              <h2 className="text-xl font-bold text-slate-800">Carga Masiva de Centros y Usuarios</h2>
              <p className="text-sm text-slate-500">Sube tu hoja de cálculo (.xlsx) para actualizar los centros y accesos.</p>
            </div>

            <div className="p-4 bg-slate-50 rounded-lg border border-slate-200 space-y-2 text-xs text-slate-600">
              <p className="font-bold text-slate-800">Estructura del Excel:</p>
              <ul className="list-disc pl-4 space-y-1">
                <li><strong>Columna A:</strong> Nombre del Centro</li>
                <li><strong>Columna B:</strong> Zona</li>
                <li><strong>Columnas C, E, G, I, K:</strong> Correo del usuario</li>
                <li><strong>Columnas D, F, H, J, L:</strong> Código Alfanumérico del usuario (Contraseña)</li>
              </ul>
            </div>

            <div className="border-2 border-dashed border-slate-300 rounded-xl p-8 text-center hover:border-blue-500 transition cursor-pointer bg-slate-50">
              <input 
                type="file" 
                accept=".xlsx, .xls"
                onChange={handleFileUploadCentres}
                className="hidden" 
                id="excel-upload" 
              />
              <label htmlFor="excel-upload" className="cursor-pointer space-y-2 block">
                <FileSpreadsheet className="w-10 h-10 text-emerald-600 mx-auto" />
                <span className="block text-sm font-medium text-slate-700">Seleccionar archivo Excel</span>
              </label>
            </div>
          </div>
        )}

        {activeTab === 'users' && currentUser.role === 'superadmin' && (
          <div className="space-y-6">
            <div>
              <h2 className="text-xl font-bold text-slate-800">Usuarios Registrados</h2>
              <p className="text-sm text-slate-500">Listado de usuarios y contraseñas (código alfanumérico).</p>
            </div>

            <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <table className="w-full text-left text-sm text-slate-600">
                <thead className="bg-slate-50 text-xs font-semibold text-slate-500 uppercase tracking-wider border-b">
                  <tr>
                    <th className="px-6 py-3">Nombre / ID</th>
                    <th className="px-6 py-3">Correo</th>
                    <th className="px-6 py-3">Código Alfanumérico</th>
                    <th className="px-6 py-3">Rol</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {users.map((u) => (
                    <tr key={u.id} className="hover:bg-slate-50">
                      <td className="px-6 py-4 font-medium text-slate-800">{u.name}</td>
                      <td className="px-6 py-4">{u.email}</td>
                      <td className="px-6 py-4 font-mono text-xs text-slate-600">{u.code}</td>
                      <td className="px-6 py-4">
                        <span className={`px-2 py-1 text-xs font-semibold rounded ${u.role === 'superadmin' ? 'bg-purple-100 text-purple-700' : 'bg-blue-100 text-blue-700'}`}>
                          {u.role}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>

      {editDocModal.open && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-xl shadow-xl max-w-md w-full p-6 space-y-4">
            <h3 className="text-lg font-bold text-slate-800">Añadir / Editar Enlace de SharePoint</h3>

            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Nombre o Título del Documento</label>
              <input 
                type="text" 
                className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                placeholder="Ej. Evaluación de Riesgos Laborales 2026"
                value={editDocModal.title}
                onChange={(e) => setEditDocModal({ ...editDocModal, title: e.target.value })}
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">URL / Enlace de SharePoint</label>
              <input 
                type="url" 
                className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                placeholder="https://empresa.sharepoint.com/..."
                value={editDocModal.link}
                onChange={(e) => setEditDocModal({ ...editDocModal, link: e.target.value })}
              />
            </div>

            <div className="flex justify-end space-x-3 pt-4 border-t">
              <button 
                onClick={() => setEditDocModal({ open: false, centreId: null, categoryKey: null, title: '', link: '' })}
                className="px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-100 rounded-lg transition"
              >
                Cancelar
              </button>
              <button 
                onClick={saveSharepointLink}
                className="px-4 py-2 text-sm font-medium bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition shadow"
              >
                Guardar Enlace
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

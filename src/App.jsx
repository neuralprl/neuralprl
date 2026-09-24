import React, { useState, useMemo } from 'react';
import { 
  Building2, 
  FileText, 
  ClipboardCheck, 
  ShieldAlert, 
  AlertTriangle, 
  CheckCircle2, 
  Clock, 
  Search, 
  Filter, 
  Eye, 
  Download, 
  Plus, 
  X, 
  UserCheck, 
  ChevronRight, 
  TrendingUp, 
  Menu, 
  Activity, 
  FileCheck,
  HardHat
} from 'lucide-react';

const INITIAL_DOCUMENTS = [
  {
    id: 'DOC-2024-001',
    title: 'Plan de Prevención de Riesgos Laborales 2024',
    category: 'Plan de Prevención',
    subCategory: 'Normativa General',
    status: 'Vigente',
    uploadDate: '2024-01-15',
    expirationDate: '2025-01-15',
    author: 'Elena Gómez (Técnico PRL)',
    fileSize: '4.2 MB',
    fileType: 'PDF',
    description: 'Documento marco con las directrices de seguridad, evaluación de riesgos generales y protocolos de evacuación.'
  },
  {
    id: 'DOC-2024-002',
    title: 'Evaluación de Riesgos - Nave Industrial B',
    category: 'Evaluación de Riesgos',
    subCategory: 'Naves Operativas',
    status: 'Pendiente Revisión',
    uploadDate: '2024-02-10',
    expirationDate: '2024-08-10',
    author: 'Carlos Ruiz (Ingeniero PRL)',
    fileSize: '8.1 MB',
    fileType: 'PDF',
    description: 'Análisis de puestos de trabajo en líneas de ensamblaje, incluyendo mapas de ruido y carga física.'
  },
  {
    id: 'DOC-2024-003',
    title: 'Entrega de EPIs - Equipo de Soldadura',
    category: 'EPIs',
    subCategory: 'Equipos de Protección',
    status: 'Vigente',
    uploadDate: '2024-03-01',
    expirationDate: '2024-09-01',
    author: 'Marta Pastor (RRHH)',
    fileSize: '1.5 MB',
    fileType: 'PDF',
    description: 'Registro firmado digitalmente de la entrega de pantallas, guantes dieléctricos y botas de seguridad.'
  },
  {
    id: 'DOC-2024-004',
    title: 'Protocolo Sanitario COVID y Gripe 2024',
    category: 'Vigilancia de la Salud',
    subCategory: 'Protocolos Médicos',
    status: 'Caducado',
    uploadDate: '2023-03-12',
    expirationDate: '2024-03-12',
    author: 'Dr. Fernando Maza (Servicio Médico)',
    fileSize: '2.0 MB',
    fileType: 'PDF',
    description: 'Normativa de actuación y recomendaciones higiénicas para el personal de oficina y almacén.'
  },
  {
    id: 'DOC-2024-005',
    title: 'Informe de Inspección de Seguridad - Carretillas',
    category: 'Inspecciones',
    subCategory: 'Maquinaria',
    status: 'Vigente',
    uploadDate: '2024-03-15',
    expirationDate: '2024-06-15',
    author: 'Carlos Ruiz (Ingeniero PRL)',
    fileSize: '3.4 MB',
    fileType: 'PDF',
    description: 'Inspección periódica del estado de frenos, señalización acústica y baterías de las carretillas elevadoras.'
  }
];

const INITIAL_CHECKLISTS = [
  {
    id: 'CHK-101',
    title: 'Inspección Semanal de Salidas de Emergencia',
    area: 'Nave A - Planta 1',
    inspector: 'Roberto Sánchez',
    date: '2024-03-20',
    itemsCount: 8,
    completedItems: 8,
    status: 'Completado',
    criticalCount: 0
  },
  {
    id: 'CHK-102',
    title: 'Revisión de Extintores y BIEs',
    area: 'Oficinas Centrales',
    inspector: 'Elena Gómez',
    date: '2024-03-22',
    itemsCount: 12,
    completedItems: 10,
    status: 'En Progreso',
    criticalCount: 1
  },
  {
    id: 'CHK-103',
    title: 'Verificación de Estado de Escaleras y Rampas',
    area: 'Zona de Carga y Descarga',
    inspector: 'David Vidal',
    date: '2024-03-25',
    itemsCount: 6,
    completedItems: 2,
    status: 'Pendiente',
    criticalCount: 0
  }
];

export default function App() {
  const [activeTab, setActiveTab] = useState('dashboard');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [documents, setDocuments] = useState(INITIAL_DOCUMENTS);
  const [checklists, setChecklists] = useState(INITIAL_CHECKLISTS);

  // Filters State
  const [searchTerm, setSearchTerm] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('Todos');
  const [statusFilter, setStatusFilter] = useState('Todos');

  // Modal State
  const [selectedDoc, setSelectedDoc] = useState(null);
  const [isNewDocModalOpen, setIsNewDocModalOpen] = useState(false);

  // New Document Form State
  const [newDocData, setNewDocData] = useState({
    title: '',
    category: 'Plan de Prevención',
    subCategory: '',
    author: '',
    expirationDate: '',
    description: ''
  });

  const stats = useMemo(() => {
    const totalDocs = documents.length;
    const activeDocs = documents.filter(d => d.status === 'Vigente').length;
    const pendingDocs = documents.filter(d => d.status === 'Pendiente Revisión').length;
    const expiredDocs = documents.filter(d => d.status === 'Caducado').length;
    const complianceRate = totalDocs > 0 ? Math.round((activeDocs / totalDocs) * 100) : 0;

    return { totalDocs, activeDocs, pendingDocs, expiredDocs, complianceRate };
  }, [documents]);

  // Filtered Documents
  const filteredDocuments = useMemo(() => {
    return documents.filter(doc => {
      const matchesSearch = doc.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
                            doc.id.toLowerCase().includes(searchTerm.toLowerCase()) ||
                            doc.author.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesCategory = categoryFilter === 'Todos' || doc.category === categoryFilter;
      const matchesStatus = statusFilter === 'Todos' || doc.status === statusFilter;
      return matchesSearch && matchesCategory && matchesStatus;
    });
  }, [documents, searchTerm, categoryFilter, statusFilter]);

  const handleCreateDocument = (e) => {
    e.preventDefault();
    if (!newDocData.title || !newDocData.author) return;

    const newDoc = {
      id: `DOC-2024-${String(documents.length + 1).padStart(3, '0')}`,
      title: newDocData.title,
      category: newDocData.category,
      subCategory: newDocData.subCategory || 'General',
      status: 'Vigente',
      uploadDate: new Date().toISOString().split('T')[0],
      expirationDate: newDocData.expirationDate || '2025-12-31',
      author: newDocData.author,
      fileSize: '2.5 MB',
      fileType: 'PDF',
      description: newDocData.description || 'Nuevo documento registrado en el sistema.'
    };

    setDocuments([newDoc, ...documents]);
    setIsNewDocModalOpen(false);
    setNewDocData({
      title: '',
      category: 'Plan de Prevención',
      subCategory: '',
      author: '',
      expirationDate: '',
      description: ''
    });
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col md:flex-row text-slate-800 font-sans">
      {}
      {/* Mobile Top Header */}
      <div className="md:hidden bg-slate-900 text-white p-4 flex justify-between items-center sticky top-0 z-40 shadow-md">
        <div className="flex items-center space-x-2">
          <HardHat className="h-7 w-7 text-amber-400" />
          <span className="font-bold text-lg tracking-wide">PRL Portal</span>
        </div>
        <button 
          onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)} 
          className="p-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200"
          aria-label="Toggle Navigation Menu"
        >
          <Menu className="h-6 w-6" />
        </button>
      </div>

      {/* Sidebar Container */}
      <aside className={`
        fixed md:static inset-y-0 left-0 z-50 w-64 bg-slate-900 text-slate-300 transform 
        ${isMobileMenuOpen ? 'translate-x-0' : '-translate-x-full'} 
        md:translate-x-0 transition-transform duration-200 ease-in-out flex flex-col justify-between shadow-xl
      `}>
        <div>
          {/* Logo & Brand */}
          <div className="p-6 hidden md:flex items-center space-x-3 border-b border-slate-800">
            <div className="p-2 bg-amber-500 rounded-xl text-slate-950 font-bold">
              <HardHat className="h-6 w-6" />
            </div>
            <div>
              <h1 className="font-bold text-white text-base tracking-tight">Prevención PRL</h1>
              <p className="text-xs text-slate-400">Gestión Documental</p>
            </div>
          </div>

          {/* Navigation Links */}
          <nav className="p-4 space-y-1">
            <button
              onClick={() => { setActiveTab('dashboard'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center space-x-3 px-4 py-3 rounded-lg font-medium text-sm transition-all ${
                activeTab === 'dashboard' 
                  ? 'bg-amber-500 text-slate-950 font-semibold shadow-sm' 
                  : 'hover:bg-slate-800 text-slate-300 hover:text-white'
              }`}
            >
              <Activity className="h-5 w-5" />
              <span>Dashboard</span>
            </button>

            <button
              onClick={() => { setActiveTab('documents'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center space-x-3 px-4 py-3 rounded-lg font-medium text-sm transition-all ${
                activeTab === 'documents' 
                  ? 'bg-amber-500 text-slate-950 font-semibold shadow-sm' 
                  : 'hover:bg-slate-800 text-slate-300 hover:text-white'
              }`}
            >
              <FileText className="h-5 w-5" />
              <span>Documentación</span>
            </button>

            <button
              onClick={() => { setActiveTab('checklists'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center space-x-3 px-4 py-3 rounded-lg font-medium text-sm transition-all ${
                activeTab === 'checklists' 
                  ? 'bg-amber-500 text-slate-950 font-semibold shadow-sm' 
                  : 'hover:bg-slate-800 text-slate-300 hover:text-white'
              }`}
            >
              <ClipboardCheck className="h-5 w-5" />
              <span>Inspecciones & Listas</span>
            </button>
          </nav>
        </div>

        {/* User Footer Profile */}
        <div className="p-4 border-t border-slate-800 bg-slate-950/40">
          <div className="flex items-center space-x-3">
            <div className="h-9 w-9 rounded-full bg-slate-700 flex items-center justify-center font-bold text-amber-400 text-sm">
              EG
            </div>
            <div className="overflow-hidden">
              <p className="text-sm font-medium text-white truncate">Elena Gómez</p>
              <p className="text-xs text-slate-400 truncate">Técnico Superior PRL</p>
            </div>
          </div>
        </div>
      </aside>

      {/* Overlay for mobile menu */}
      {isMobileMenuOpen && (
        <div 
          onClick={() => setIsMobileMenuOpen(false)} 
          className="fixed inset-0 bg-slate-900/60 z-40 md:hidden backdrop-blur-sm"
        />
      )}

      {}
      <main className="flex-1 overflow-y-auto p-4 md:p-8">
        {/* Header Title Bar */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-8">
          <div>
            <h2 className="text-2xl font-bold text-slate-900 tracking-tight">
              {activeTab === 'dashboard' && 'Panel de Control de Seguridad laboral'}
              {activeTab === 'documents' && 'Gestión de Documentos PRL'}
              {activeTab === 'checklists' && 'Inspecciones y Verificaciones de Campo'}
            </h2>
            <p className="text-slate-500 text-sm mt-1">
              Monitoreo en tiempo real, auditorías y cumplimiento normativo de prevención.
            </p>
          </div>

          {activeTab === 'documents' && (
            <button
              onClick={() => setIsNewDocModalOpen(true)}
              className="inline-flex items-center justify-center space-x-2 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2.5 rounded-lg text-sm font-semibold shadow-sm transition-colors focus:ring-2 focus:ring-emerald-500 focus:ring-offset-2"
            >
              <Plus className="h-4 w-4" />
              <span>Nuevo Documento</span>
            </button>
          )}
        </div>

        {}
        {activeTab === 'dashboard' && (
          <div className="space-y-8">
            {/* KPI Cards */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Documentos Totales</span>
                  <div className="p-2 bg-blue-50 text-blue-600 rounded-lg">
                    <FileText className="h-5 w-5" />
                  </div>
                </div>
                <div className="mt-4 flex items-baseline justify-between">
                  <span className="text-2xl font-bold text-slate-900">{stats.totalDocs}</span>
                  <span className="text-xs font-medium text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded-full flex items-center">
                    <TrendingUp className="h-3 w-3 mr-1" /> Activos
                  </span>
                </div>
              </div>

              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Nivel de Cumplimiento</span>
                  <div className="p-2 bg-emerald-50 text-emerald-600 rounded-lg">
                    <CheckCircle2 className="h-5 w-5" />
                  </div>
                </div>
                <div className="mt-4 flex items-baseline justify-between">
                  <span className="text-2xl font-bold text-slate-900">{stats.complianceRate}%</span>
                  <span className="text-xs text-slate-500">Vigencia reglamentaria</span>
                </div>
              </div>

              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Pendientes de Revisión</span>
                  <div className="p-2 bg-amber-50 text-amber-600 rounded-lg">
                    <Clock className="h-5 w-5" />
                  </div>
                </div>
                <div className="mt-4 flex items-baseline justify-between">
                  <span className="text-2xl font-bold text-slate-900">{stats.pendingDocs}</span>
                  <span className="text-xs text-amber-700 bg-amber-50 px-2 py-0.5 rounded-full">Atención requerida</span>
                </div>
              </div>

              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">Documentos Caducados</span>
                  <div className="p-2 bg-rose-50 text-rose-600 rounded-lg">
                    <AlertTriangle className="h-5 w-5" />
                  </div>
                </div>
                <div className="mt-4 flex items-baseline justify-between">
                  <span className="text-2xl font-bold text-slate-900">{stats.expiredDocs}</span>
                  <span className="text-xs text-rose-700 bg-rose-50 px-2 py-0.5 rounded-full">Riesgo legal</span>
                </div>
              </div>
            </div>

            {/* Recent Documents & Quick Checklists Overview */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
              {/* Recent Activity Table */}
              <div className="lg:col-span-2 bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between mb-4">
                  <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
                    <FileCheck className="h-5 w-5 text-slate-600" />
                    Documentación Reciente
                  </h3>
                  <button 
                    onClick={() => setActiveTab('documents')} 
                    className="text-xs font-semibold text-blue-600 hover:text-blue-800 flex items-center"
                  >
                    Ver todos <ChevronRight className="h-4 w-4 ml-0.5" />
                  </button>
                </div>
                
                <div className="divide-y divide-slate-100">
                  {documents.slice(0, 4).map((doc) => (
                    <div key={doc.id} className="py-3 flex items-center justify-between hover:bg-slate-50 px-2 rounded-lg transition-colors">
                      <div className="flex items-center space-x-3 overflow-hidden pr-2">
                        <div className="p-2 bg-slate-100 text-slate-600 rounded-lg shrink-0">
                          <FileText className="h-5 w-5" />
                        </div>
                        <div className="truncate">
                          <p className="text-sm font-semibold text-slate-800 truncate">{doc.title}</p>
                          <p className="text-xs text-slate-500">{doc.category} • {doc.uploadDate}</p>
                        </div>
                      </div>
                      <span className={`text-xs px-2.5 py-1 rounded-full font-medium shrink-0 ${
                        doc.status === 'Vigente' ? 'bg-emerald-100 text-emerald-800' :
                        doc.status === 'Pendiente Revisión' ? 'bg-amber-100 text-amber-800' :
                        'bg-rose-100 text-rose-800'
                      }`}>
                        {doc.status}
                      </span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Checklists Summary Box */}
              <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex flex-col justify-between">
                <div>
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
                      <ClipboardCheck className="h-5 w-5 text-slate-600" />
                      Estado de Inspecciones
                    </h3>
                  </div>

                  <div className="space-y-4">
                    {checklists.map((chk) => (
                      <div key={chk.id} className="p-3 bg-slate-50 rounded-lg border border-slate-100">
                        <div className="flex justify-between items-start mb-2">
                          <h4 className="text-xs font-bold text-slate-800">{chk.title}</h4>
                          <span className={`text-[10px] px-2 py-0.5 rounded font-semibold ${
                            chk.status === 'Completado' ? 'bg-emerald-100 text-emerald-700' :
                            chk.status === 'En Progreso' ? 'bg-amber-100 text-amber-700' :
                            'bg-slate-200 text-slate-700'
                          }`}>
                            {chk.status}
                          </span>
                        </div>
                        <div className="w-full bg-slate-200 rounded-full h-1.5 mb-1">
                          <div 
                            className="bg-blue-600 h-1.5 rounded-full" 
                            style={{ width: `${(chk.completedItems / chk.itemsCount) * 100}%` }}
                          />
                        </div>
                        <p className="text-[11px] text-slate-500 text-right">{chk.completedItems}/{chk.itemsCount} Ítems</p>
                      </div>
                    ))}
                  </div>
                </div>

                <button
                  onClick={() => setActiveTab('checklists')}
                  className="mt-6 w-full py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-lg transition-colors text-center"
                >
                  Ir al módulo de Inspecciones
                </button>
              </div>
            </div>
          </div>
        )}

        {}
        {activeTab === 'documents' && (
          <div className="space-y-6">
            {/* Filter Bar */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col md:flex-row gap-4 justify-between items-stretch md:items-center">
              {/* Search input */}
              <div className="relative flex-1">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-slate-400" />
                <input
                  type="text"
                  placeholder="Buscar documento por título, código o responsable..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-slate-50/50"
                />
              </div>

              {/* Filters */}
              <div className="flex flex-wrap gap-2 items-center">
                <div className="flex items-center space-x-2 bg-slate-50 px-3 py-2 border border-slate-200 rounded-lg text-sm">
                  <Filter className="h-4 w-4 text-slate-500" />
                  <select
                    value={categoryFilter}
                    onChange={(e) => setCategoryFilter(e.target.value)}
                    className="bg-transparent focus:outline-none text-slate-700 text-sm"
                  >
                    <option value="Todos">Todas las Categorías</option>
                    <option value="Plan de Prevención">Plan de Prevención</option>
                    <option value="Evaluación de Riesgos">Evaluación de Riesgos</option>
                    <option value="EPIs">EPIs</option>
                    <option value="Vigilancia de la Salud">Vigilancia de la Salud</option>
                    <option value="Inspecciones">Inspecciones</option>
                  </select>
                </div>

                <div className="flex items-center space-x-2 bg-slate-50 px-3 py-2 border border-slate-200 rounded-lg text-sm">
                  <select
                    value={statusFilter}
                    onChange={(e) => setStatusFilter(e.target.value)}
                    className="bg-transparent focus:outline-none text-slate-700 text-sm"
                  >
                    <option value="Todos">Todos los Estados</option>
                    <option value="Vigente">Vigente</option>
                    <option value="Pendiente Revisión">Pendiente Revisión</option>
                    <option value="Caducado">Caducado</option>
                  </select>
                </div>
              </div>
            </div>

            {/* Documents Table */}
            <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left text-sm text-slate-600">
                  <thead className="bg-slate-50 border-b border-slate-200 text-slate-700 font-semibold text-xs uppercase tracking-wider">
                    <tr>
                      <th className="px-6 py-3.5">Código / Título</th>
                      <th className="px-6 py-3.5">Categoría</th>
                      <th className="px-6 py-3.5">Responsable</th>
                      <th className="px-6 py-3.5">Fecha Caducidad</th>
                      <th className="px-6 py-3.5">Estado</th>
                      <th className="px-6 py-3.5 text-right">Acciones</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {filteredDocuments.length > 0 ? (
                      filteredDocuments.map((doc) => (
                        <tr key={doc.id} className="hover:bg-slate-50/80 transition-colors">
                          <td className="px-6 py-4">
                            <div className="font-bold text-slate-900">{doc.title}</div>
                            <div className="text-xs text-slate-400">{doc.id} • {doc.fileSize}</div>
                          </td>
                          <td className="px-6 py-4">
                            <span className="bg-slate-100 text-slate-700 px-2.5 py-1 rounded-md text-xs font-medium border border-slate-200">
                              {doc.category}
                            </span>
                          </td>
                          <td className="px-6 py-4 text-xs text-slate-600">
                            {doc.author}
                          </td>
                          <td className="px-6 py-4 text-xs font-medium text-slate-700">
                            {doc.expirationDate}
                          </td>
                          <td className="px-6 py-4">
                            <span className={`inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium ${
                              doc.status === 'Vigente' ? 'bg-emerald-100 text-emerald-800' :
                              doc.status === 'Pendiente Revisión' ? 'bg-amber-100 text-amber-800' :
                              'bg-rose-100 text-rose-800'
                            }`}>
                              {doc.status}
                            </span>
                          </td>
                          <td className="px-6 py-4 text-right space-x-2">
                            <button
                              onClick={() => setSelectedDoc(doc)}
                              className="p-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg transition-colors"
                              title="Ver detalle"
                            >
                              <Eye className="h-4 w-4" />
                            </button>
                            <button
                              onClick={() => alert(`Descargando el archivo: ${doc.title}`)}
                              className="p-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg transition-colors"
                              title="Descargar"
                            >
                              <Download className="h-4 w-4" />
                            </button>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan="6" className="text-center py-12 text-slate-400">
                          <AlertTriangle className="h-8 w-8 mx-auto mb-2 text-slate-300" />
                          No se encontraron documentos que coincidan con los filtros.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {}
        {activeTab === 'checklists' && (
          <div className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              {checklists.map((chk) => (
                <div key={chk.id} className="bg-white rounded-xl border border-slate-200 shadow-sm p-5 flex flex-col justify-between">
                  <div>
                    <div className="flex justify-between items-start mb-3">
                      <span className="text-xs font-mono bg-slate-100 text-slate-600 px-2 py-0.5 rounded">
                        {chk.id}
                      </span>
                      <span className={`text-xs px-2.5 py-0.5 rounded-full font-semibold ${
                        chk.status === 'Completado' ? 'bg-emerald-100 text-emerald-700' :
                        chk.status === 'En Progreso' ? 'bg-amber-100 text-amber-700' :
                        'bg-slate-100 text-slate-600'
                      }`}>
                        {chk.status}
                      </span>
                    </div>

                    <h3 className="font-bold text-slate-900 text-base mb-2">{chk.title}</h3>
                    
                    <div className="space-y-1.5 text-xs text-slate-500 mb-4">
                      <p className="flex items-center"><Building2 className="h-3.5 w-3.5 mr-1.5" /> {chk.area}</p>
                      <p className="flex items-center"><UserCheck className="h-3.5 w-3.5 mr-1.5" /> Inspector: {chk.inspector}</p>
                      <p className="flex items-center"><Clock className="h-3.5 w-3.5 mr-1.5" /> Fecha: {chk.date}</p>
                    </div>
                  </div>

                  <div>
                    <div className="flex justify-between text-xs font-medium text-slate-600 mb-1">
                      <span>Progreso de verificación</span>
                      <span>{chk.completedItems} / {chk.itemsCount}</span>
                    </div>
                    <div className="w-full bg-slate-100 rounded-full h-2 mb-4">
                      <div 
                        className="bg-blue-600 h-2 rounded-full transition-all duration-300"
                        style={{ width: `${(chk.completedItems / chk.itemsCount) * 100}%` }}
                      />
                    </div>

                    <button
                      onClick={() => alert(`Abriendo cuestionario interactivo para ${chk.id}`)}
                      className="w-full py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold rounded-lg transition-colors"
                    >
                      Realizar Inspección
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </main>

      {}
      {selectedDoc && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl border border-slate-100 relative animate-in fade-in zoom-in-95 duration-150">
            <button
              onClick={() => setSelectedDoc(null)}
              className="absolute right-4 top-4 text-slate-400 hover:text-slate-600 p-1 rounded-full hover:bg-slate-100"
            >
              <X className="h-5 w-5" />
            </button>

            <div className="flex items-center space-x-3 mb-4">
              <div className="p-3 bg-blue-50 text-blue-600 rounded-xl">
                <FileText className="h-6 w-6" />
              </div>
              <div>
                <span className="text-xs font-mono text-slate-400">{selectedDoc.id}</span>
                <h3 className="text-xl font-bold text-slate-900">{selectedDoc.title}</h3>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-4 py-4 border-y border-slate-100 text-sm">
              <div>
                <span className="text-xs text-slate-400 block mb-0.5">Categoría</span>
                <span className="font-semibold text-slate-700">{selectedDoc.category}</span>
              </div>
              <div>
                <span className="text-xs text-slate-400 block mb-0.5">Estado de Vigencia</span>
                <span className={`inline-block px-2 py-0.5 rounded text-xs font-semibold ${
                  selectedDoc.status === 'Vigente' ? 'bg-emerald-100 text-emerald-800' :
                  selectedDoc.status === 'Pendiente Revisión' ? 'bg-amber-100 text-amber-800' :
                  'bg-rose-100 text-rose-800'
                }`}>
                  {selectedDoc.status}
                </span>
              </div>
              <div>
                <span className="text-xs text-slate-400 block mb-0.5">Responsable / Autor</span>
                <span className="font-medium text-slate-700">{selectedDoc.author}</span>
              </div>
              <div>
                <span className="text-xs text-slate-400 block mb-0.5">Fecha de Vencimiento</span>
                <span className="font-medium text-slate-700">{selectedDoc.expirationDate}</span>
              </div>
            </div>

            <div className="mt-4">
              <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-1">Descripción del Archivo</h4>
              <p className="text-sm text-slate-600 bg-slate-50 p-3 rounded-lg border border-slate-100">
                {selectedDoc.description}
              </p>
            </div>

            <div className="mt-6 flex justify-end space-x-3">
              <button
                onClick={() => setSelectedDoc(null)}
                className="px-4 py-2 border border-slate-200 text-slate-600 hover:bg-slate-50 rounded-lg text-sm font-medium"
              >
                Cerrar
              </button>
              <button
                onClick={() => alert(`Descargando ${selectedDoc.title}...`)}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-sm font-semibold flex items-center space-x-2"
              >
                <Download className="h-4 w-4" />
                <span>Descargar PDF</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {}
      {isNewDocModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-lg w-full p-6 shadow-2xl border border-slate-100 relative">
            <button
              onClick={() => setIsNewDocModalOpen(false)}
              className="absolute right-4 top-4 text-slate-400 hover:text-slate-600 p-1 rounded-full hover:bg-slate-100"
            >
              <X className="h-5 w-5" />
            </button>

            <h3 className="text-lg font-bold text-slate-900 mb-4">Registrar Nuevo Documento PRL</h3>

            <form onSubmit={handleCreateDocument} className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Título del Documento *</label>
                <input
                  type="text"
                  required
                  placeholder="Ej: Evaluación Ergonomica 2024"
                  value={newDocData.title}
                  onChange={(e) => setNewDocData({ ...newDocData, title: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Categoría</label>
                  <select
                    value={newDocData.category}
                    onChange={(e) => setNewDocData({ ...newDocData, category: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none bg-white"
                  >
                    <option value="Plan de Prevención">Plan de Prevención</option>
                    <option value="Evaluación de Riesgos">Evaluación de Riesgos</option>
                    <option value="EPIs">EPIs</option>
                    <option value="Vigilancia de la Salud">Vigilancia de la Salud</option>
                    <option value="Inspecciones">Inspecciones</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Fecha de Caducidad</label>
                  <input
                    type="date"
                    value={newDocData.expirationDate}
                    onChange={(e) => setNewDocData({ ...newDocData, expirationDate: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none"
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Técnico / Responsable *</label>
                <input
                  type="text"
                  required
                  placeholder="Ej: Laura Martínez"
                  value={newDocData.author}
                  onChange={(e) => setNewDocData({ ...newDocData, author: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Descripción</label>
                <textarea
                  rows="3"
                  placeholder="Resumen o alcance del documento..."
                  value={newDocData.description}
                  onChange={(e) => setNewDocData({ ...newDocData, description: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              <div className="pt-2 flex justify-end space-x-3">
                <button
                  type="button"
                  onClick={() => setIsNewDocModalOpen(false)}
                  className="px-4 py-2 border border-slate-200 text-slate-600 hover:bg-slate-50 rounded-lg text-sm font-medium"
                >
                  Cancelar
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-sm font-semibold shadow-sm"
                >
                  Guardar Documento
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

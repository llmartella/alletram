import React, { useState, useEffect } from 'react';
import { Plus, Save, Trash2, Download, ArrowLeft, CheckCircle, XCircle, Edit2 } from 'lucide-react';

const QualificationTestApp = () => {
  const [scenarios, setScenarios] = useState([]);
  const [currentView, setCurrentView] = useState('list');
  const [currentScenario, setCurrentScenario] = useState(null);
  const [editingId, setEditingId] = useState(null);

  useEffect(() => {
    loadScenarios();
  }, []);

  const loadScenarios = async () => {
    try {
      const result = await window.storage.list('scenario:');
      if (result && result.keys) {
        const loadedScenarios = await Promise.all(
          result.keys.map(async (key) => {
            try {
              const data = await window.storage.get(key);
              return data ? JSON.parse(data.value) : null;
            } catch {
              return null;
            }
          })
        );
        setScenarios(loadedScenarios.filter(s => s !== null));
      }
    } catch (error) {
      console.error('Error loading scenarios:', error);
    }
  };

  const saveScenario = async (scenario) => {
    try {
      const scenarioToSave = {
        ...scenario,
        id: scenario.id || `scenario_${Date.now()}`,
        lastModified: new Date().toISOString()
      };
      
      await window.storage.set(
        `scenario:${scenarioToSave.id}`,
        JSON.stringify(scenarioToSave)
      );
      
      await loadScenarios();
      setCurrentView('list');
      setCurrentScenario(null);
      setEditingId(null);
    } catch (error) {
      console.error('Error saving scenario:', error);
      alert('Failed to save scenario');
    }
  };

  const deleteScenario = async (id) => {
    if (confirm('Are you sure you want to delete this scenario?')) {
      try {
        await window.storage.delete(`scenario:${id}`);
        await loadScenarios();
      } catch (error) {
        console.error('Error deleting scenario:', error);
      }
    }
  };

  const startNewScenario = () => {
    const newScenario = {
      name: '',
      testType: 'yoy',
      qualifyingSkus: '',
      earningSkus: '',
      currentYearEarningRate: '',
      targetPercent: '',
      numPriorYears: 1,
      priorYearTimeframes: [
        { label: 'Prior Year 1', startDate: '', endDate: '' }
      ],
      currentYearTimeframe: { startDate: '', endDate: '' },
      minQualificationPercent: '',
      earningPercent: '',
      tests: [],
      nextTestNumber: 1
    };
    setCurrentScenario(newScenario);
    setCurrentView('create');
    setEditingId(null);
  };

  const editScenario = (scenario) => {
    // Ensure tests array and nextTestNumber exist
    const scenarioToEdit = {
      ...scenario,
      tests: scenario.tests || [],
      nextTestNumber: scenario.nextTestNumber || 1
    };
    setCurrentScenario(scenarioToEdit);
    setCurrentView('edit');
    setEditingId(scenario.id);
  };

  if (currentView === 'list') {
    return <ScenarioList 
      scenarios={scenarios}
      onNew={startNewScenario}
      onEdit={editScenario}
      onDelete={deleteScenario}
    />;
  }

  return (
    <ScenarioEditor
      scenario={currentScenario}
      isEditing={currentView === 'edit'}
      onSave={saveScenario}
      onCancel={() => {
        setCurrentView('list');
        setCurrentScenario(null);
        setEditingId(null);
      }}
      onChange={setCurrentScenario}
    />
  );
};

const ScenarioList = ({ scenarios, onNew, onEdit, onDelete }) => {
  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-6xl mx-auto">
        <div className="flex justify-between items-center mb-6">
          <h1 className="text-3xl font-bold text-gray-900">Qualification Test Scenarios</h1>
          <button
            onClick={onNew}
            className="flex items-center gap-2 bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 transition-colors"
          >
            <Plus size={20} />
            New Scenario
          </button>
        </div>

        {scenarios.length === 0 ? (
          <div className="bg-white rounded-lg shadow p-12 text-center">
            <p className="text-gray-500 mb-4">No scenarios created yet</p>
            <button
              onClick={onNew}
              className="text-blue-600 hover:text-blue-700 font-medium"
            >
              Create your first scenario
            </button>
          </div>
        ) : (
          <div className="grid gap-4">
            {scenarios.map((scenario) => (
              <div key={scenario.id} className="bg-white rounded-lg shadow p-6 hover:shadow-md transition-shadow">
                <div className="flex justify-between items-start">
                  <div className="flex-1">
                    <h3 className="text-xl font-semibold text-gray-900 mb-2">{scenario.name}</h3>
                    <div className="grid grid-cols-2 gap-4 text-sm text-gray-600 mb-3">
                      <div>
                        <span className="font-medium">Test Type:</span>{' '}
                        {scenario.testType === 'yoy' ? 'Year over Year' :
                         scenario.testType === 'minQual' ? 'Min Qualification + Earnings' :
                         'Earnings Only'}
                      </div>
                      <div>
                        <span className="font-medium">Tests:</span> {scenario.tests?.length || 0}
                      </div>
                    </div>
                    {scenario.tests && scenario.tests.length > 0 && (
                      <div className="text-xs text-gray-500">
                        Qualified: {scenario.tests.filter(t => t.results?.qualified).length} | 
                        Not Qualified: {scenario.tests.filter(t => !t.results?.qualified).length}
                      </div>
                    )}
                    <div className="mt-2 text-xs text-gray-400">
                      Last modified: {new Date(scenario.lastModified).toLocaleString()}
                    </div>
                  </div>
                  <div className="flex gap-2 ml-4">
                    <button
                      onClick={() => onEdit(scenario)}
                      className="px-3 py-1 bg-gray-100 text-gray-700 rounded hover:bg-gray-200 transition-colors"
                    >
                      Edit
                    </button>
                    <button
                      onClick={() => onDelete(scenario.id)}
                      className="px-3 py-1 bg-red-100 text-red-700 rounded hover:bg-red-200 transition-colors"
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

const ScenarioEditor = ({ scenario, isEditing, onSave, onCancel, onChange }) => {
  const [activeTab, setActiveTab] = useState('criteria');
  const [selectedTestId, setSelectedTestId] = useState(null);

  const updateField = (field, value) => {
    onChange({ ...scenario, [field]: value });
  };

  const updatePriorYearTimeframe = (index, field, value) => {
    const newTimeframes = [...scenario.priorYearTimeframes];
    newTimeframes[index] = { ...newTimeframes[index], [field]: value };
    updateField('priorYearTimeframes', newTimeframes);
  };

  const updateCurrentYearTimeframe = (field, value) => {
    const newTimeframe = { ...scenario.currentYearTimeframe, [field]: value };
    updateField('currentYearTimeframe', newTimeframe);
  };

  const updateNumPriorYears = (num) => {
    const numPriorYears = parseInt(num);
    const newTimeframes = [];
    
    for (let i = 0; i < numPriorYears; i++) {
      newTimeframes.push(
        scenario.priorYearTimeframes[i] || { 
          label: `Prior Year ${i + 1}`,
          startDate: '', 
          endDate: '' 
        }
      );
    }
    
    updateField('numPriorYears', numPriorYears);
    updateField('priorYearTimeframes', newTimeframes);
  };

  const addTest = () => {
    const currentTests = scenario.tests || [];
    const currentNextTestNumber = scenario.nextTestNumber || 1;
    const testId = String(currentNextTestNumber).padStart(6, '0');
    const newTest = {
      id: testId,
      label: `Test ${currentNextTestNumber}`,
      transactions: [],
      results: null
    };
    onChange({
      ...scenario,
      tests: [...currentTests, newTest],
      nextTestNumber: currentNextTestNumber + 1
    });
    setSelectedTestId(testId);
  };

  const deleteTest = (testId) => {
    if (confirm('Are you sure you want to delete this test?')) {
      const currentTests = scenario.tests || [];
      const newTests = currentTests.filter(t => t.id !== testId);
      onChange({
        ...scenario,
        tests: newTests
      });
      if (selectedTestId === testId) {
        setSelectedTestId(null);
      }
    }
  };

  const updateTest = (testId, updatedTest) => {
    const currentTests = scenario.tests || [];
    const newTests = currentTests.map(t => 
      t.id === testId ? updatedTest : t
    );
    onChange({
      ...scenario,
      tests: newTests
    });
  };

  const handleSave = () => {
    if (!scenario.name.trim()) {
      alert('Please enter a scenario name');
      return;
    }
    
    // Recalculate all test results before saving
    const updatedTests = scenario.tests.map(test => ({
      ...test,
      results: performCalculations(scenario, test)
    }));
    
    onSave({ ...scenario, tests: updatedTests });
  };

  const exportData = () => {
    const dataStr = JSON.stringify(scenario, null, 2);
    const dataBlob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(dataBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${scenario.name.replace(/\s+/g, '_')}_${Date.now()}.json`;
    link.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-6xl mx-auto">
        <div className="flex justify-between items-center mb-6">
          <div className="flex items-center gap-4">
            <button
              onClick={onCancel}
              className="p-2 hover:bg-gray-200 rounded-lg transition-colors"
            >
              <ArrowLeft size={24} />
            </button>
            <h1 className="text-3xl font-bold text-gray-900">
              {isEditing ? 'Edit Scenario' : 'New Scenario'}
            </h1>
          </div>
          <div className="flex gap-2">
            <button
              onClick={exportData}
              className="flex items-center gap-2 px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition-colors"
            >
              <Download size={20} />
              Export
            </button>
            <button
              onClick={handleSave}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
            >
              <Save size={20} />
              Save
            </button>
          </div>
        </div>

        <div className="bg-white rounded-lg shadow mb-6">
          <div className="border-b border-gray-200">
            <nav className="flex">
              <button
                onClick={() => setActiveTab('criteria')}
                className={`px-6 py-3 font-medium ${
                  activeTab === 'criteria'
                    ? 'border-b-2 border-blue-600 text-blue-600'
                    : 'text-gray-600 hover:text-gray-900'
                }`}
              >
                Criteria
              </button>
              <button
                onClick={() => setActiveTab('tests')}
                className={`px-6 py-3 font-medium ${
                  activeTab === 'tests'
                    ? 'border-b-2 border-blue-600 text-blue-600'
                    : 'text-gray-600 hover:text-gray-900'
                }`}
              >
                Tests ({scenario.tests?.length || 0})
              </button>
            </nav>
          </div>

          <div className="p-6">
            {activeTab === 'criteria' && (
              <CriteriaForm
                scenario={scenario}
                updateField={updateField}
                updatePriorYearTimeframe={updatePriorYearTimeframe}
                updateCurrentYearTimeframe={updateCurrentYearTimeframe}
                updateNumPriorYears={updateNumPriorYears}
              />
            )}

            {activeTab === 'tests' && (
              <TestsManager
                scenario={scenario}
                tests={scenario.tests || []}
                selectedTestId={selectedTestId}
                onSelectTest={setSelectedTestId}
                onAddTest={addTest}
                onDeleteTest={deleteTest}
                onUpdateTest={updateTest}
              />
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

const CriteriaForm = ({ scenario, updateField, updatePriorYearTimeframe, updateCurrentYearTimeframe, updateNumPriorYears }) => {
  return (
    <div className="space-y-6">
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          Scenario Name *
        </label>
        <input
          type="text"
          value={scenario.name}
          onChange={(e) => updateField('name', e.target.value)}
          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          placeholder="e.g., Q4 2024 Customer Analysis"
        />
      </div>

      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          Test Type *
        </label>
        <select
          value={scenario.testType}
          onChange={(e) => updateField('testType', e.target.value)}
          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
        >
          <option value="yoy">Year over Year (YoY)</option>
          <option value="minQual">Current Year Minimum Qualification + Earnings</option>
          <option value="earningsOnly">Earnings Only</option>
        </select>
      </div>

      <div className="grid grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Qualifying SKUs (comma-separated)
          </label>
          <input
            type="text"
            value={scenario.qualifyingSkus}
            onChange={(e) => updateField('qualifyingSkus', e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
            placeholder="e.g., SKU001, SKU002, SKU003"
          />
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Earning SKUs (comma-separated)
          </label>
          <input
            type="text"
            value={scenario.earningSkus}
            onChange={(e) => updateField('earningSkus', e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
            placeholder="e.g., SKU001, SKU002"
          />
        </div>
      </div>

      {scenario.testType === 'yoy' && (
        <>
          <div className="border-t border-gray-300 pt-6 mt-6">
            <h3 className="text-lg font-semibold text-gray-900 mb-4">Prior Year(s) Definition</h3>
            
            <div className="grid grid-cols-2 gap-4 mb-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Target % (for qualification)
                </label>
                <input
                  type="number"
                  step="0.01"
                  value={scenario.targetPercent}
                  onChange={(e) => updateField('targetPercent', e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  placeholder="e.g., 95"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Number of Prior Years
                </label>
                <input
                  type="number"
                  min="1"
                  max="10"
                  value={scenario.numPriorYears}
                  onChange={(e) => updateNumPriorYears(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-3">
                Prior Year Timeframes
              </label>
              <div className="space-y-3">
                {scenario.priorYearTimeframes.map((tf, index) => (
                  <div key={index} className="flex items-center gap-4 p-4 bg-gray-50 rounded-lg">
                    <span className="font-medium text-gray-700 w-32">{tf.label}</span>
                    <input
                      type="date"
                      value={tf.startDate}
                      onChange={(e) => updatePriorYearTimeframe(index, 'startDate', e.target.value)}
                      className="flex-1 px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    />
                    <span className="text-gray-500">to</span>
                    <input
                      type="date"
                      value={tf.endDate}
                      onChange={(e) => updatePriorYearTimeframe(index, 'endDate', e.target.value)}
                      className="flex-1 px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    />
                  </div>
                ))}
              </div>
            </div>
          </div>

          <div className="border-t border-gray-300 pt-6 mt-6">
            <h3 className="text-lg font-semibold text-gray-900 mb-4">Current Year Definition</h3>
            
            <div className="mb-4">
              <label className="block text-sm font-medium text-gray-700 mb-3">
                Current Year Timeframe
              </label>
              <div className="flex items-center gap-4 p-4 bg-blue-50 rounded-lg">
                <span className="font-medium text-gray-700 w-32">Current Year</span>
                <input
                  type="date"
                  value={scenario.currentYearTimeframe.startDate}
                  onChange={(e) => updateCurrentYearTimeframe('startDate', e.target.value)}
                  className="flex-1 px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
                <span className="text-gray-500">to</span>
                <input
                  type="date"
                  value={scenario.currentYearTimeframe.endDate}
                  onChange={(e) => updateCurrentYearTimeframe('endDate', e.target.value)}
                  className="flex-1 px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Current Year Earning Rate (%)
              </label>
              <input
                type="number"
                step="0.01"
                value={scenario.currentYearEarningRate}
                onChange={(e) => updateField('currentYearEarningRate', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                placeholder="e.g., 5.5"
              />
            </div>
          </div>
        </>
      )}

      {scenario.testType === 'minQual' && (
        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Minimum Qualification % *
            </label>
            <input
              type="number"
              step="0.01"
              value={scenario.minQualificationPercent}
              onChange={(e) => updateField('minQualificationPercent', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              placeholder="e.g., 80"
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Earning % *
            </label>
            <input
              type="number"
              step="0.01"
              value={scenario.earningPercent}
              onChange={(e) => updateField('earningPercent', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              placeholder="e.g., 5"
            />
          </div>
        </div>
      )}

      {scenario.testType === 'earningsOnly' && (
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Earning % *
          </label>
          <input
            type="number"
            step="0.01"
            value={scenario.earningPercent}
            onChange={(e) => updateField('earningPercent', e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
            placeholder="e.g., 5"
          />
        </div>
      )}
    </div>
  );
};

const TestsManager = ({ scenario, tests, selectedTestId, onSelectTest, onAddTest, onDeleteTest, onUpdateTest }) => {
  const selectedTest = tests.find(t => t.id === selectedTestId);

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <h3 className="text-lg font-semibold text-gray-900">Tests</h3>
        <button
          onClick={onAddTest}
          className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
        >
          <Plus size={20} />
          Add Test
        </button>
      </div>

      {tests.length === 0 ? (
        <div className="text-center py-12 text-gray-500">
          No tests created yet. Click "Add Test" to begin.
        </div>
      ) : (
        <div className="grid grid-cols-3 gap-6">
          <div className="col-span-1 space-y-2">
            <h4 className="text-sm font-medium text-gray-700 mb-2">Test List</h4>
            {tests.map((test) => (
              <div
                key={test.id}
                onClick={() => onSelectTest(test.id)}
                className={`p-3 rounded-lg cursor-pointer transition-colors ${
                  selectedTestId === test.id
                    ? 'bg-blue-100 border-2 border-blue-500'
                    : 'bg-gray-50 border-2 border-transparent hover:bg-gray-100'
                }`}
              >
                <div className="flex justify-between items-start mb-2">
                  <div>
                    <div className="font-mono text-xs text-gray-500">#{test.id}</div>
                    <div className="font-medium text-gray-900">{test.label}</div>
                  </div>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      onDeleteTest(test.id);
                    }}
                    className="text-red-600 hover:text-red-700"
                  >
                    <Trash2 size={14} />
                  </button>
                </div>
                <div className="text-xs text-gray-600 mb-2">
                  {test.transactions.length} transaction{test.transactions.length !== 1 ? 's' : ''}
                </div>
                {test.results && (
                  <div className="space-y-1">
                    <div className="flex items-center gap-1">
                      {test.results.qualified ? (
                        <span className="text-xs text-green-600 flex items-center gap-1">
                          <CheckCircle size={12} /> Qualified
                        </span>
                      ) : (
                        <span className="text-xs text-red-600 flex items-center gap-1">
                          <XCircle size={12} /> Not Qualified
                        </span>
                      )}
                    </div>
                    <div className="text-xs font-semibold text-blue-700">
                      {test.results.qualified ? `${test.results.earnings?.toFixed(2) || '0.00'}` : '--'}
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>

          <div className="col-span-2">
            {selectedTest ? (
              <TestEditor
                scenario={scenario}
                test={selectedTest}
                onUpdate={(updatedTest) => onUpdateTest(selectedTest.id, updatedTest)}
              />
            ) : (
              <div className="text-center py-12 text-gray-500">
                Select a test to view transactions and results
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

const TestEditor = ({ scenario, test, onUpdate }) => {
  const addTransaction = () => {
    const newTransaction = {
      id: `txn_${Date.now()}`,
      date: '',
      sku: '',
      quantity: '',
      price: '',
      total: ''
    };
    const updatedTest = {
      ...test,
      transactions: [...test.transactions, newTransaction]
    };
    onUpdate(updatedTest);
  };

  const updateTransaction = (index, field, value) => {
    const newTransactions = [...test.transactions];
    newTransactions[index] = { ...newTransactions[index], [field]: value };
    
    if (field === 'quantity' || field === 'price') {
      const quantity = parseFloat(field === 'quantity' ? value : newTransactions[index].quantity) || 0;
      const price = parseFloat(field === 'price' ? value : newTransactions[index].price) || 0;
      newTransactions[index].total = (quantity * price).toFixed(2);
    }
    
    onUpdate({ ...test, transactions: newTransactions });
  };

  const deleteTransaction = (index) => {
    const newTransactions = test.transactions.filter((_, i) => i !== index);
    onUpdate({ ...test, transactions: newTransactions });
  };

  const recalculate = () => {
    const results = performCalculations(scenario, test);
    onUpdate({ ...test, results });
  };

  const handleKeyDown = (e, index, field) => {
    if (e.key === 'Tab' && !e.shiftKey) {
      // If we're on the last field (total) of the last transaction, add a new one
      if (field === 'price' && index === test.transactions.length - 1) {
        e.preventDefault();
        addTransaction();
        // Focus will naturally move to the new row's first field
        setTimeout(() => {
          const inputs = document.querySelectorAll('input[type="date"]');
          const lastInput = inputs[inputs.length - 1];
          if (lastInput) lastInput.focus();
        }, 50);
      }
    }
  };

  const results = test.results || performCalculations(scenario, test);

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center">
        <div>
          <div className="font-mono text-sm text-gray-500">Test ID: #{test.id}</div>
          <h4 className="text-lg font-semibold text-gray-900">{test.label}</h4>
        </div>
      </div>

      <div className="grid grid-cols-2 gap-4">
        <div className="p-4 bg-gray-50 rounded-lg">
          <h5 className="text-sm font-medium text-gray-700 mb-2">Qualification Status</h5>
          <div className="flex items-center gap-2">
            {results.qualified ? (
              <>
                <CheckCircle size={20} className="text-green-600" />
                <span className="text-lg font-bold text-green-600">Qualified</span>
              </>
            ) : (
              <>
                <XCircle size={20} className="text-red-600" />
                <span className="text-lg font-bold text-red-600">Not Qualified</span>
              </>
            )}
          </div>
        </div>

        <div className="p-4 bg-gray-50 rounded-lg">
          <h5 className="text-sm font-medium text-gray-700 mb-2">Total Earnings</h5>
          <div className="text-lg font-bold text-blue-600">
            {results.qualified ? `${results.earnings?.toFixed(2) || '0.00'}` : '--'}
          </div>
        </div>
      </div>

      {scenario.testType === 'yoy' && results.details && (
        <div className="p-4 bg-blue-50 rounded-lg border border-blue-200">
          <h5 className="text-sm font-medium text-blue-900 mb-3">Year over Year Details</h5>
          <div className="space-y-2 text-sm">
            {results.details.yearTotals && results.details.yearTotals.map((year, index) => (
              <div key={index} className="flex justify-between">
                <span className="text-blue-800">{year.label}:</span>
                <span className="font-medium text-blue-900">${year.total.toFixed(2)}</span>
              </div>
            ))}
            {results.details.achievedPercent !== undefined && (
              <>
                <div className="border-t border-blue-300 my-2"></div>
                <div className="flex justify-between">
                  <span className="text-blue-800">Achievement Rate:</span>
                  <span className="font-medium text-blue-900">{results.details.achievedPercent.toFixed(2)}%</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-blue-800">Target Rate:</span>
                  <span className="font-medium text-blue-900">{scenario.targetPercent}%</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-blue-800">Earning Rate:</span>
                  <span className="font-medium text-blue-900">{scenario.currentYearEarningRate}%</span>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      <div>
        <div className="flex justify-between items-center mb-3">
          <h5 className="text-sm font-medium text-gray-700">Transactions</h5>
          <div className="flex gap-2">
            <button
              onClick={addTransaction}
              className="flex items-center gap-1 px-3 py-1 bg-green-600 text-white text-sm rounded hover:bg-green-700 transition-colors"
            >
              <Plus size={16} />
              Add Transaction
            </button>
            <button
              onClick={recalculate}
              className="px-3 py-1 bg-blue-600 text-white text-sm rounded hover:bg-blue-700 transition-colors"
            >
              Recalculate
            </button>
          </div>
        </div>

        {test.transactions.length === 0 ? (
          <div className="text-center py-8 text-gray-500 text-sm">
            No transactions yet. Click "Add Transaction" to begin.
          </div>
        ) : (
          <div className="space-y-2">
            {test.transactions.map((txn, index) => (
              <div key={txn.id} className="p-3 bg-white border border-gray-200 rounded-lg">
                <div className="grid grid-cols-6 gap-2">
                  <div>
                    <label className="block text-xs font-medium text-gray-600 mb-1">Date</label>
                    <input
                      type="date"
                      value={txn.date}
                      onChange={(e) => updateTransaction(index, 'date', e.target.value)}
                      onKeyDown={(e) => handleKeyDown(e, index, 'date')}
                      className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-gray-600 mb-1">SKU</label>
                    <input
                      type="text"
                      value={txn.sku}
                      onChange={(e) => updateTransaction(index, 'sku', e.target.value)}
                      onKeyDown={(e) => handleKeyDown(e, index, 'sku')}
                      className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-gray-600 mb-1">Quantity</label>
                    <input
                      type="number"
                      step="0.01"
                      value={txn.quantity}
                      onChange={(e) => updateTransaction(index, 'quantity', e.target.value)}
                      onKeyDown={(e) => handleKeyDown(e, index, 'quantity')}
                      className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-gray-600 mb-1">Price</label>
                    <input
                      type="number"
                      step="0.01"
                      value={txn.price}
                      onChange={(e) => updateTransaction(index, 'price', e.target.value)}
                      onKeyDown={(e) => handleKeyDown(e, index, 'price')}
                      className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-gray-600 mb-1">Total</label>
                    <input
                      type="text"
                      value={txn.total}
                      readOnly
                      className="w-full px-2 py-1 text-sm border border-gray-300 rounded bg-gray-50"
                    />
                  </div>
                  <div className="flex items-end">
                    <button
                      onClick={() => deleteTransaction(index)}
                      className="w-full px-2 py-1 bg-red-100 text-red-700 rounded hover:bg-red-200 transition-colors"
                    >
                      <Trash2 size={14} className="mx-auto" />
                    </button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

const performCalculations = (scenario, test) => {
  const qualifyingSkus = scenario.qualifyingSkus.split(',').map(s => s.trim()).filter(Boolean);
  const earningSkus = scenario.earningSkus.split(',').map(s => s.trim()).filter(Boolean);
  
  let qualified = false;
  let earnings = 0;
  let details = {};

  if (scenario.testType === 'yoy') {
    const yearTotals = scenario.priorYearTimeframes.map(tf => {
      const total = test.transactions
        .filter(txn => {
          const txnDate = new Date(txn.date);
          const start = new Date(tf.startDate);
          const end = new Date(tf.endDate);
          return txnDate >= start && txnDate <= end && qualifyingSkus.includes(txn.sku);
        })
        .reduce((sum, txn) => sum + parseFloat(txn.total || 0), 0);
      
      return { label: tf.label, total };
    });

    const currentYearTotal = test.transactions
      .filter(txn => {
        const txnDate = new Date(txn.date);
        const start = new Date(scenario.currentYearTimeframe.startDate);
        const end = new Date(scenario.currentYearTimeframe.endDate);
        return txnDate >= start && txnDate <= end && qualifyingSkus.includes(txn.sku);
      })
      .reduce((sum, txn) => sum + parseFloat(txn.total || 0), 0);

    const mostRecentPriorYearTotal = yearTotals[yearTotals.length - 1]?.total || 0;
    
    const achievedPercent = mostRecentPriorYearTotal > 0 
      ? (currentYearTotal / mostRecentPriorYearTotal) * 100 
      : 0;
    
    const targetPercent = parseFloat(scenario.targetPercent) || 95;
    qualified = achievedPercent >= targetPercent;

    if (qualified) {
      earnings = test.transactions
        .filter(txn => {
          const txnDate = new Date(txn.date);
          const start = new Date(scenario.currentYearTimeframe.startDate);
          const end = new Date(scenario.currentYearTimeframe.endDate);
          return txnDate >= start && txnDate <= end && earningSkus.includes(txn.sku);
        })
        .reduce((sum, txn) => sum + parseFloat(txn.total || 0), 0) 
        * (parseFloat(scenario.currentYearEarningRate) || 0) / 100;
    }

    yearTotals.push({ label: 'Current Year', total: currentYearTotal });

    details = { 
      yearTotals, 
      achievedPercent,
      qualifyingTransactions: test.transactions.filter(txn => qualifyingSkus.includes(txn.sku)).length,
      earningTransactions: test.transactions.filter(txn => earningSkus.includes(txn.sku)).length
    };

  } else if (scenario.testType === 'minQual') {
    qualified = true;
    
    earnings = test.transactions
      .filter(txn => earningSkus.includes(txn.sku))
      .reduce((sum, txn) => sum + parseFloat(txn.total || 0), 0)
      * (parseFloat(scenario.earningPercent) || 0) / 100;

    details = {
      qualifyingTransactions: test.transactions.filter(txn => qualifyingSkus.includes(txn.sku)).length,
      earningTransactions: test.transactions.filter(txn => earningSkus.includes(txn.sku)).length
    };

  } else if (scenario.testType === 'earningsOnly') {
    qualified = true;
    
    earnings = test.transactions
      .filter(txn => earningSkus.includes(txn.sku))
      .reduce((sum, txn) => sum + parseFloat(txn.total || 0), 0)
      * (parseFloat(scenario.earningPercent) || 0) / 100;

    details = {
      qualifyingTransactions: 0,
      earningTransactions: test.transactions.filter(txn => earningSkus.includes(txn.sku)).length
    };
  }

  return { qualified, earnings, details };
};

export default QualificationTestApp;

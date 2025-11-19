document.addEventListener('DOMContentLoaded', () => {
    // --- Logika dla dynamicznego dodawania pól mapowania (admin) ---
    const addCustomFieldBtn = document.getElementById('add-custom-field-btn');
    if (addCustomFieldBtn) {
        const customFieldsContainer = document.getElementById('custom-fields-container');
        let customFieldIndex = customFieldsContainer.children.length;

        addCustomFieldBtn.addEventListener('click', () => {
            const newFieldDiv = document.createElement('div');
            newFieldDiv.classList.add('form-group', 'custom-field-group', 'form-row', 'align-items-center');
            newFieldDiv.innerHTML = `
                <div class="col-sm-8">
                     <div class="custom-field-inputs">
                        <input type="text" name="custom_name_${customFieldIndex}" class="form-control" placeholder="Nazwa własnego pola" required>
                        <input type="text" name="custom_cell_${customFieldIndex}" class="form-control" placeholder="Adres komórki" required>
                    </div>
                </div>
                <div class="col-sm-3">
                     <div class="form-check">
                        <input class="form-check-input" type="checkbox" name="custom_is_required_${customFieldIndex}" id="custom_is_required_${customFieldIndex}">
                        <label class="form-check-label" for="custom_is_required_${customFieldIndex}">Wymagane</label>
                    </div>
                </div>
                <div class="col-sm-1">
                    <button type="button" class="btn btn-warning btn-sm remove-field-btn">Usuń</button>
                </div>
            `;
            customFieldsContainer.appendChild(newFieldDiv);
            newFieldDiv.querySelector('.remove-field-btn').addEventListener('click', (e) => e.target.closest('.custom-field-group').remove());
            customFieldIndex++;
        });

        document.querySelectorAll('.remove-field-btn').forEach(btn => {
            btn.addEventListener('click', (e) => e.target.closest('.custom-field-group').remove());
        });
    }

    // --- Logika dla formularza generowania certyfikatów (user) ---
    const certForm = document.getElementById('cert-form');
    if (certForm) {
        const templateSelect = document.getElementById('template_id');
        const generatorCustomFieldsContainer = document.getElementById('generator-custom-fields');

        function updateFormFields() {
            const selectedTemplateId = templateSelect.value;

            document.querySelectorAll('.dynamic-field').forEach(el => el.style.display = 'none');
            generatorCustomFieldsContainer.innerHTML = '';

            if (!selectedTemplateId) return;

            fetch(`/api/template/${selectedTemplateId}`)
                .then(response => {
                    if (!response.ok) throw new Error('Nie udało się pobrać szczegółów szablonu.');
                    return response.json();
                })
                .then(data => {
                    const mappedFields = data.mapped_fields || [];
                    const numberingFormat = data.numbering_format || '';

                    const startNumberGroup = document.getElementById('field-group-numer_certyfikatu');
                    if (startNumberGroup && (numberingFormat.includes('{NUMER}') || numberingFormat.includes('{NUMER_STARTOWY}'))) {
                        startNumberGroup.style.display = 'block';
                    }

                    mappedFields.forEach(field => {
                        const fieldName = field.name;
                        const isRequired = field.required;

                        const group = document.getElementById(`field-group-${fieldName}`);
                        if (group) {
                            group.style.display = 'block';
                            const input = group.querySelector('input, select');
                            if (input) {
                                input.required = isRequired;
                                if (input.tagName === 'INPUT' && input.type !== 'date' && input.type !== 'number') {
                                    input.placeholder = isRequired ? '' : 'opcjonalne';
                                }
                            }
                        }
                        if (fieldName === 'waznosc') {
                           const container = document.getElementById('validity-value-container');
                           if(container) container.style.display = 'block';
                           document.getElementById('validity_type').dispatchEvent(new Event('change'));
                        }

                        if (fieldName.startsWith('custom_')) {
                            const cleanName = fieldName.replace('custom_', '').replace(/_/g, ' ');
                            const formKey = fieldName.replace('custom_', '');

                            const newFieldDiv = document.createElement('div');
                            newFieldDiv.classList.add('form-group');
                            newFieldDiv.innerHTML = `
                                <label for="custom_field_${formKey}">${cleanName.charAt(0).toUpperCase() + cleanName.slice(1)}:</label>
                                <input type="text" id="custom_field_${formKey}" name="custom_field_${formKey}" class="form-control"
                                       ${isRequired ? 'required' : ''}
                                       placeholder="${isRequired ? '' : 'opcjonalne'}">
                            `;
                            generatorCustomFieldsContainer.appendChild(newFieldDiv);
                        }
                    });
                })
                .catch(error => {
                    console.error(error);
                    const statusDiv = document.getElementById('status');
                    if(statusDiv) statusDiv.textContent = error.message;
                });
        }

        templateSelect.addEventListener('change', updateFormFields);
        if (templateSelect.value) {
            updateFormFields();
        }

        // PRZYWRÓCONA, PEŁNA LOGIKA OBSŁUGI WAŻNOŚCI
        const validityTypeSelect = document.getElementById('validity_type');
        if (validityTypeSelect) {
            validityTypeSelect.addEventListener('change', function() {
                const validityYearsGroup = document.getElementById('validity-years-group');
                const validityDateGroup = document.getElementById('validity-date-group');
                if (this.value === 'years') {
                    validityYearsGroup.style.display = 'block';
                    validityDateGroup.style.display = 'none';
                } else if (this.value === 'date') {
                    validityYearsGroup.style.display = 'none';
                    validityDateGroup.style.display = 'block';
                } else {
                    validityYearsGroup.style.display = 'none';
                    validityDateGroup.style.display = 'none';
                }
            });
        }

        // PRZYWRÓCONA, PEŁNA LOGIKA OBSŁUGI WYSYŁKI FORMULARZA
        certForm.addEventListener('submit', (event) => {
            event.preventDefault();
            const generateBtn = certForm.querySelector('#generate-btn');
            const statusDiv = document.getElementById('status');
            generateBtn.disabled = true;
            generateBtn.textContent = 'Generowanie...';
            statusDiv.className = 'alert alert-info';
            statusDiv.textContent = 'Przygotowywanie danych, proszę czekać...';

            fetch('/api/generate', { method: 'POST', body: new FormData(certForm) })
            .then(response => {
                if (!response.ok) return response.json().then(err => { throw new Error(err.error || 'Wystąpił nieznany błąd serwera.') });
                return response.blob();
            })
            .then(blob => {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = 'certyfikaty.zip';
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                a.remove();
                statusDiv.className = 'alert alert-success';
                statusDiv.textContent = 'Gotowe! Plik certyfikaty.zip został pobrany.';
            })
            .catch(error => {
                statusDiv.className = 'alert alert-danger';
                statusDiv.textContent = `Błąd: ${error.message}`;
            })
            .finally(() => {
                generateBtn.disabled = false;
                generateBtn.textContent = 'Generuj certyfikaty';
            });
        });
    }

    // PRZYWRÓCONA LOGIKA POTWIERDZENIA USUNIĘCIA
    document.querySelectorAll('.delete-form').forEach(form => {
        form.addEventListener('submit', function(event) {
            if (!confirm(`Czy na pewno chcesz usunąć szablon "${this.dataset.templateName}"? Tej operacji nie można cofnąć.`)) {
                event.preventDefault();
            }
        });
    });
});
fetch('../data/bhargav.CSV.xls')
    .then(response => response.arrayBuffer())
    .then(data => {
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        createDashboard(jsonData);
    })
    .catch(error => console.error('Error loading Excel file:', error));

function createDashboard(data) {
    // Extract data for visualizations
    const courseNames = [...new Set(data.map(d => d['Name of the Cource']))];
    const punctuality = courseNames.map(name => average(data, name, 'Punctuality'));
    const overallSatisfaction = courseNames.map(name => average(data, name, 'Overall Satisfaction'));
    const radarData = averageMetrics(data, courseNames[0]); // Radar chart for the first course
    const pieData = calculatePieData(data, 'Overall Satisfaction');
    const timeSeriesData = calculateTimeSeries(data, 'Overall Satisfaction');

    // Render charts
    renderBarChart('barChart', courseNames, punctuality, 'Punctuality');
    renderRadarChart('radarChart', radarData, courseNames[0]);
    renderPieChart('pieChart', pieData);
    renderLineChart('lineChart', timeSeriesData.timestamps, timeSeriesData.values);
}

function average(data, course, column) {
    const courseData = data.filter(d => d['Name of the Cource'] === course);
    const sum = courseData.reduce((acc, curr) => acc + (curr[column] || 0), 0);
    return sum / courseData.length;
}

function averageMetrics(data, course) {
    const metrics = ['Punctuality', 'Teaching Methodology', 'Course Content Organization',
                     'Knowledge about the subject', 'Doubt Clarification', 'Interaction with students'];
    return metrics.map(metric => average(data, course, metric));
}

function calculatePieData(data, column) {
    const counts = {};
    data.forEach(d => {
        const value = d[column];
        counts[value] = (counts[value] || 0) + 1;
    });
    return counts;
}

function calculateTimeSeries(data, column) {
    const timestamps = [...new Set(data.map(d => d.Timestamp))].sort();
    const values = timestamps.map(ts => {
        const filtered = data.filter(d => d.Timestamp === ts);
        return filtered.reduce((sum, d) => sum + (d[column] || 0), 0) / filtered.length;
    });
    return { timestamps, values };
}

function renderBarChart(canvasId, labels, data, label) {
    new Chart(document.getElementById(canvasId).getContext('2d'), {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: label,
                data: data,
                backgroundColor: '#3498db'
            }]
        },
        options: {
            responsive: true,
            plugins: {
                title: { display: true, text: `${label} Across Courses` }
            }
        }
    });
}

function renderRadarChart(canvasId, data, courseName) {
    new Chart(document.getElementById(canvasId).getContext('2d'), {
        type: 'radar',
        data: {
            labels: ['Punctuality', 'Teaching Methodology', 'Content Organization', 
                     'Knowledge', 'Doubt Clarification', 'Interaction'],
            datasets: [{
                label: `Metrics for ${courseName}`,
                data: data,
                backgroundColor: 'rgba(52, 152, 219, 0.2)',
                borderColor: '#3498db',
                pointBackgroundColor: '#3498db'
            }]
        },
        options: {
            responsive: true,
            plugins: {
                title: { display: true, text: `Metrics for ${courseName}` }
            }
        }
    });
}

function renderPieChart(canvasId, data) {
    new Chart(document.getElementById(canvasId).getContext('2d'), {
        type: 'pie',
        data: {
            labels: Object.keys(data),
            datasets: [{
                label: 'Overall Satisfaction Distribution',
                data: Object.values(data),
                backgroundColor: ['#3498db', '#2ecc71', '#e74c3c', '#f1c40f', '#9b59b6']
            }]
        },
        options: {
            responsive: true,
            plugins: {
                title: { display: true, text: 'Overall Satisfaction Distribution' }
            }
        }
    });
}

function renderLineChart(canvasId, labels, data) {
    new Chart(document.getElementById(canvasId).getContext('2d'), {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Overall Satisfaction Over Time',
                data: data,
                borderColor: '#3498db',
                backgroundColor: 'rgba(52, 152, 219, 0.2)',
                fill: true
            }]
        },
        options: {
            responsive: true,
            plugins: {
                title: { display: true, text: 'Overall Satisfaction Over Time' }
            },
            scales: {
                x: { title: { display: true, text: 'Timestamp' } },
                y: { title: { display: true, text: 'Average Rating (1 to 5)' }, min: 0, max: 5 }
            }
        }
    });
}

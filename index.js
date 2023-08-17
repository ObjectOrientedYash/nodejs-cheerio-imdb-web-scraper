const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');

const addTitlesToExcel = async (moviePayload) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Top movies');
  worksheet.addRow(['Movie Title', 'Movie Rating']);
  moviePayload.forEach((movie) => {
    worksheet.addRow([movie.title, movie.rating]);
  });
  worksheet.getColumn('A').alignment = { horizontal: 'center' };
  worksheet.getColumn('B').alignment = { horizontal: 'center' };
  worksheet.getColumn('A').width = 50;
  worksheet.getColumn('B').width = 50;
  await workbook.xlsx.writeFile('movies.xlsx');
  console.log('Movies added to Excel sheet successfully.');
};
const fetchTopMovies = async () => {
  try {
    const moviePayload = [];
    const response = await axios.get(
      'https://www.imdb.com/chart/top/?ref_=nv_mv_250'
    );
    const html = response.data;
    const $ = cheerio.load(html); 
    $('tbody.lister-list tr').each((index, ele) => {
      const movie = $(ele);
      const title = movie.find('td.titleColumn a').text().trim();
      const rating = movie.find('td.ratingColumn.imdbRating').text().trim();
      console.log(title);
      console.log(rating);
      const movieData = {
        title,
        rating,
      };
      moviePayload.push(movieData);
    });
    return moviePayload;
  } catch (err) {
    console.log(err);
  }
};

fetchTopMovies().then((moviePayload) => {
  addTitlesToExcel(moviePayload);
});

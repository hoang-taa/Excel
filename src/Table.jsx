/* eslint-disable react/prop-types */
// eslint-disable-next-line react/prop-types
function TableData({ data = [] }) {
  return (
    <div className='flex flex-col '>
      <div className='-m-1.5 overflow-x-auto'>
        <div className='p-1.5 min-w-full inline-block align-middle'>
          <div className='border overflow-hidden dark:border-gray-700'>
            <table className='min-w-full divide-y divide-gray-200 dark:divide-gray-700'>
              <thead>
                <tr className='divide-x divide-gray-200 dark:divide-gray-700'>
                  <th
                    scope='col'
                    className='px-6 py-3 text-start text-xs font-medium text-gray-500 uppercase'
                  >
                    Tên site
                  </th>
                  <th
                    scope='col'
                    className='px-6 py-3 text-start text-xs font-medium text-gray-500 uppercase'
                  >
                    Số lượng & chủng loại UBBP
                  </th>
                  <th
                    scope='col'
                    className='px-6 py-3 text-start text-xs font-medium text-gray-500 uppercase'
                  >
                    Số lượng và chủng loại RRU
                  </th>
                  <th
                    scope='col'
                    className='px-6 py-3 text-start text-xs font-medium text-gray-500 uppercase'
                  >
                    Số cell hiện tại
                  </th>
                  <th
                    scope='col'
                    className='px-6 py-3 text-start text-xs font-medium text-gray-500 uppercase'
                  >
                    Số cell hỗ trợ tối đa
                  </th>
                  <th
                    scope='col'
                    className='px-6 py-3 text-start text-xs font-medium text-gray-500 uppercase'
                  >
                    Đáp ứng Moran (Y/N)
                  </th>
                </tr>
              </thead>
              <tbody className='divide-y divide-gray-200 dark:divide-gray-700'>
                {data.map((item, index) => {
                  return (
                    <tr key={index}>
                      <td className='px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-800 dark:text-gray-200'>
                        {item["Tên site"]}
                      </td>
                      <td className='px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-800 dark:text-gray-200'>
                        {item["Số lượng & chủng loại UBBP"]}
                      </td>
                      <td className='px-6 py-4 whitespace-pre-line text-sm font-medium text-gray-800 dark:text-gray-200 max-w-28'>
                        {item["Số lượng và chủng loại RRU"]}
                      </td>
                      <td className='px-6 py-4 whitespace-pre-line text-sm font-medium text-gray-800 dark:text-gray-200 max-w-28'>
                        {item["Số cell hiện tại"]}
                      </td>
                      <td className='px-6 py-4 whitespace-pre-line text-sm font-medium text-gray-800 dark:text-gray-200 max-w-28'>
                        {item["Số cell hỗ trợ tối đa"]}
                      </td>
                      <td className='px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-800 dark:text-gray-200'>
                        {item["Đáp ứng Moran (Y/N)"]}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

export default TableData;
